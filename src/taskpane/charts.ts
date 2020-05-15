export async function addChart(context): Promise<void> {
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
        chart.format.fill.setSolidColor("white");
  
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