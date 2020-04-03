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
          const dataByCountry = await getCovidDataByCountry(countries[i][0]);

          // If valid data was found add on to countryData for subsequent table and chart creation
          if (dataByCountry.length > 0) {
            ranegeData.push([countries[i].toString()]);
            countryData.push(dataByCountry);
          } else {
            ranegeData.push([`${countries[i]} is not a real country name, silly!`]);
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

async function getCovidDataByCountry(country: string): Promise<any> {
  return new Promise<object>(async (resolve, reject) => {
    try {
      // Get data for country
      let databyCountry: Object[] = [];

      const dataByCountryApiUrl = `https://covid2019-api.herokuapp.com/v2/current`;
      const response = await fetch(dataByCountryApiUrl);
      const text = await response.text();
      const countryData = JSON.parse(text);

      for (let j = 0; j < countryData.data.length; j++) {
        if (countryData.data[j].location.toLowerCase() == country.toLowerCase()) {
          databyCountry.push(countryData.data[j]);
          break;
        }
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
      var covidTable = sheet.tables.add("A1:E1", true /*hasHeaders*/);
      covidTable.name = "CovidTable";
      covidTable.getHeaderRowRange().values = [["Country", "Active", "ComfirmedCases", "Recovered", "Deaths"]];

      // Add rows to table
      for (let i = 0; i < data.length; i++) {
        covidTable.rows.add(null /*add rows to the end of the table*/, [
          [data[i][0].location, data[i][0].active, data[i][0].confirmed, data[i][0].recovered, data[i][0].deaths],
        ]);
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
      chart.legend.position = "left"
      chart.legend.format.fill.setSolidColor("white");      

      // Don't add data labels if too much data - makes chart look cluttered
      if (rowCount < 5) {
        chart.dataLabels.format.font.size = 12;
        chart.dataLabels.textOrientation = 90;
        chart.dataLabels.format.font.color = "black";
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