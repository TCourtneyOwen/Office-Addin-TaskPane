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
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("values");
      await context.sync();
      const country = range.values[0][0];

      if (country !== "") {
        const dataByCountry = await getCovidDataByCountry(country);
        await addTableForCountry(dataByCountry, context);
        await addChart(context);
      }

      logMessgae(range.address);
    });
  } catch (error) {
    console.error(error);
  }
}

function logMessgae(message) {
  console.log(message);
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

async function getCovidDataByCountry(country: string): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    const serverResponse: any = {};
    try {
      // Get data for country and add to databyCountry
      const dataByCountryApiUrl = `https://covid2019-api.herokuapp.com/country/${country}`;
      const response = await fetch(dataByCountryApiUrl);
      serverResponse["status"] = response.status;
      const text = await response.text();
      resolve(JSON.parse(text));
    } catch (err) {
      serverResponse["status"] = err;
      reject(serverResponse);
    }
  });
}

async function addTable(data: any, context) {
  var sheet = context.workbook.worksheets.getItem("Sheet1");
  var covidTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
  covidTable.name = "CovidTable";
  covidTable.getHeaderRowRange().values = [["Country", "ComfirmedCases", "Recovered", "Deaths"]];

  const countries = Object.values(data);
  for (var key in countries) {
    const country = countries[key];
    const countryData = Object.getOwnPropertyNames(country)

    covidTable.rows.add(null /*add rows to the end of the table*/, [
      [countryData[0], country[countryData[0]].confirmed, country[countryData[0]].recovered, country[countryData[0]].deaths],
    ]);
  }
}

async function addTableForCountry(data: any, context): Promise<string> {
  return new Promise<string>(async (resolve, reject) => {
    try {
      // const currentTable = context.worksheets.worksheets.getItem("Sheet1").tables.getItem("CovidTable");
      // if (currentTable !== undefined) {
      //   currentTable.delete();
      // }
      
      var sheet = context.workbook.worksheets.getItem("Sheet1");
      var covidTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      covidTable.name = "CovidTable";
      covidTable.getHeaderRowRange().values = [["Country", "ComfirmedCases", "Recovered", "Deaths"]];

      const countryData = Object.getOwnPropertyNames(data);

      covidTable.rows.add(null /*add rows to the end of the table*/, [
        [countryData[0], _.get(data, `${countryData[0]}.confirmed`), _.get(data, `${countryData[0]}.recovered`), _.get(data, `${countryData[0]}.deaths`)],
      ]);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      const tableRange = covidTable.getRange();
      tableRange.load('address');
      sheet.activate();
      await context.sync();
      resolve(tableRange.address);
    } catch (err) {
      reject(err);
    }
  });
}

async function addChart(context): Promise<void> {
  return new Promise<void>(async (resolve, reject) => {
    try {
      // const currentChart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("Covid19 Data");
      // if (currentChart !== undefined) {
      //   currentChart.delete();
      //   context.sync();
      // }
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      var dataRange = sheet.getRange("A1:D2");
      var chart = sheet.charts.add("3DColumnStacked", dataRange, "auto");

      chart.title.text = "Covid19 Data";
      chart.legend.position = "right"
      chart.legend.format.fill.setSolidColor("white");
      resolve();
    } catch (err) {
      reject(err);
    }
  });
}