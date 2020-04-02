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
      let ranegeData = [];
      let countryData = [];

      for (let i = 0; i < countries.length; i++) {
        if (countries[i].toString() === "") {
          ranegeData.push(["No country name entered"]);
        } else {
          const dataByCountry = await getCovidDataByCountry(countries[i]);
          
          // Check to see if valid data was actually returned. If not, country specified was invalid
          let found = false;
          for (let [key] of Object.entries(dataByCountry[0])) {
            if (key.toLowerCase() === countries[i][0].toLowerCase()) {
              found = true;
              break;
            }
          }

          // If valid data was found add on to countryData for subsequent table and chart creation
          if (found) {
            ranegeData.push([countries[i].toString()]);
            countryData.push(dataByCountry);
          } else {
            ranegeData.push([`${countries[i]} is not a valid country name`]);
          }
        }
      }

      // Create table and chart if valid country data was gathered
      if (countryData.length > 0) {
        await addTableForCountry(countryData, context);
        await addChart(context);
      }

      // Update cells to inform user if empty cell or invalid country was entered
      range.values = ranegeData;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function getCovidDataByCountry(country: any[]): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    const serverResponse: any = {};
    try {
      // Get data for country[ies]
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
      // Delete table if it currently exists
      const currentTable = context.workbook.worksheets.getActiveWorksheet().tables.getItemOrNullObject("CovidTable");
      if (currentTable) {
        currentTable.delete();
        await context.sync();
      }

      // Add table to worksheet
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var covidTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      covidTable.name = "CovidTable";
      covidTable.getHeaderRowRange().values = [["Country", "ComfirmedCases", "Recovered", "Deaths"]];

      // Add rows to table
      for (let i = 0; i < data.length; i++) {
        for (var key in data[i][0]) {
          if (key === "dt" || key == "ts") {
            continue;
          }
          const country = key;
          const countryData = data[i][0][key];
  
          if (countryData[0] !== "dt") {
            covidTable.rows.add(null /*add rows to the end of the table*/, [
              [country, countryData.confirmed, countryData.recovered, countryData.deaths],
            ]);
          }
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
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const currentChart = sheet.charts.getItemOrNullObject("Covid19Chart");
      if (currentChart) {
        currentChart.delete();
        await context.sync();
      }

      // Get table range
      const currentTable = context.workbook.worksheets.getActiveWorksheet().tables.getItemOrNullObject("CovidTable");
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