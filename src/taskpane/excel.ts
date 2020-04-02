/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as fetch from "isomorphic-fetch";
import * as _ from "lodash";
/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async context => {

      const range = context.workbook.getSelectedRange();

      // Load range values
      range.load("values");
      await context.sync();
      const countries = range.values;

      for (let i = 0; i < countries.length; i++) {
        if (countries[i].toString() === "") {
          range.values[i] = [[`No country named entered`]];
        } else {
          const dataByCountry = await getCovidDataByCountry(countries[i]);
          const country = countries[i].toString().toLowerCase();
          const found = dataByCountry[Object.keys(dataByCountry).find(key => key.toLowerCase() === country.toLowerCase())];

          if (found) {
            await addTableForCountry(dataByCountry, context);
            await addChart(context);
          } else {
            range.values[i] = [[`${countries[i]} is not a valid country name`]];
          }
        }
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function getCovidData(): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    const serverResponse: any = {};
    try {

      const apiUrl: string = `https://covid2019-api.herokuapp.com/v2/total`;
      const response = await fetch(apiUrl);
      serverResponse["status"] = response.status;
      const text = await response.text();
      resolve(JSON.parse(text));
    } catch (err) {
      serverResponse["status"] = err;
      reject(serverResponse);
    }
  });
}

async function getCovidDataForAllCountries(): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    const serverResponse: any = {};
    try {
      const countriesApiUrl: string = `https://covid2019-api.herokuapp.com/countries`;
      const response = await fetch(countriesApiUrl);
      serverResponse["status"] = response.status;
      const text = await response.text();
      const countries: any = JSON.parse(text);
      let databyCountry: Object[] = []

      for (let i = 0; i < 10; i++) {
        // Get data for country and add to databyCountry
        const countryData: any = await getCovidDataByCountry(countries.countries[i]);
        databyCountry.push(countryData);
      }

      resolve(databyCountry);
    } catch (err) {
      serverResponse["status"] = err;
      reject(serverResponse);
    }
  });
}

async function getCovidDataByCountry(country: any[][]): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    const serverResponse: any = {};
    try {
      // Get data for country and add to databyCountry
      let databyCountry: Object[] = []
      for (let i = 0; i < country.length; i++) {
        const dataByCountryApiUrl = `https://covid2019-api.herokuapp.com/country/${country[i]}`;
        const response = await fetch(dataByCountryApiUrl);
        serverResponse["status"] = response.status;
        const text = await response.text();
        const countryData = JSON.parse(text);
        databyCountry.push(countryData);
      }
      resolve(databyCountry);
    } catch (err) {
      reject(err);
    }
  });
}

async function addTableForCountry(data: any, context): Promise<void | string> {
  return new Promise<void>(async (resolve, reject) => {
    try {
      const currentTable = context.workbook.worksheets.getItem("Sheet1").tables.getItemOrNullObject("CovidTable");
      if (currentTable) {
        currentTable.delete();
        await context.sync();
      }

      var sheet = context.workbook.worksheets.getItem("Sheet1");
      var covidTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      covidTable.name = "CovidTable";
      covidTable.getHeaderRowRange().values = [["Country", "ComfirmedCases", "Recovered", "Deaths"]];

      const countries = Object.values(data);
      for (var key in countries) {
        const country = countries[key];
        const countryData = Object.getOwnPropertyNames(country);

        if (countryData[0] !== "dt") {
          covidTable.rows.add(null /*add rows to the end of the table*/, [
            [countryData[0], country[countryData[0]].confirmed, country[countryData[0]].recovered, country[countryData[0]].deaths],
          ]);
        }
      }

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      sheet.activate();
      await context.sync();
      resolve();
    } catch (err) {
      reject(err);
    }
  });
}

async function addChart(context): Promise<void> {
  return new Promise<void>(async (resolve, reject) => {
    try {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const currentChart = context.workbook.worksheets.getItem("Sheet1").charts.getItemOrNullObject("Covid19Chart");
      if (currentChart) {
        currentChart.delete();
        await context.sync();
      }

      // Get table range
      const currentTable = context.workbook.worksheets.getItem("Sheet1").tables.getItemOrNullObject("CovidTable");
      const range = currentTable.getRange();
      range.load("address");
      await context.sync();
      const updatedRange = range.address.slice(range.address.indexOf("!") + 1);
      const dataRange = sheet.getRange(updatedRange);

      // Get row count
      const rows = currentTable.rows;
      rows.load("count");
      await context.sync();
      const rowCount = rows.count;      

      let chart = sheet.charts.add(rowCount < 5 ? "3DColumnClustered" : "3DColumnStacked", dataRange, "auto");
      chart.name = "Covid19Chart";
      chart.title.text = "COVID-19 Data";
      chart.legend.position = "right"
      chart.legend.format.fill.setSolidColor("white");

      // Don't add data labels if too much data - makes chart look cluttered
      if (rowCount < 5) {
        chart.dataLabels.format.font.size = 12;
        chart.dataLabels.format.font.color = "black";
        chart.dataLabels.textOrientation = 90;
      }

      chart.height = 300;
      chart.width = 500;

      await context.sync();
      resolve();
    } catch (err) {
      reject(err);
    }
  });
}