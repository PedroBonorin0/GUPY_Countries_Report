const express = require("express");
const axios = require('axios');
const app = express();

// =========Excel4Node decumentation steps=========

// Require library
var xl = require('excel4node');

app.get('/', async (req, res) => {
  // Create a new instance of a Workbook class
  var wb = new xl.Workbook();

  // Add Worksheet to the workbook
  var ws = wb.addWorksheet('Countries List');

  // Create a reusable style
  var basicStyle = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
    },
    numberFormat: '#,##0.00',
  });

  var titleStyle = wb.createStyle({
    font: {
      color: '#4F4F4F',
      size: 16,
      bold: true
    },
    alignment: {
      horizontal: "center",
    }
  });

  var headerStyle = wb.createStyle({
    font: {
      color: '#808080',
      size: 12,
      bold: true
    },
  });

  // Write Title
  ws.cell(1, 1, 1, 4, true)
    .string("Countries List")
    .style(titleStyle);

  // Write Headers
  ws.cell(2, 1)
    .string("Name")
    .style(headerStyle);

  ws.cell(2, 2)
    .string("Capital")
    .style(headerStyle);

  ws.cell(2, 3)
    .string("Area")
    .style(headerStyle);

  ws.cell(2, 4)
    .string("Currencies")
    .style(headerStyle);

  try {
    const response = await axios.get("https://restcountries.com/v3.1/all");
    
    const allCities = Object.values(response.data);

    // Order allCities By Name
    allCities.sort((a, b) => {
      if (a.name.common > b.name.common)
          return 1;
      if (a.name.common < b.name.common)
          return -1;
      return 0;
    });
    
    // Make Row on Xlsx for every country
    allCities.forEach((country) => {
      // Write Name
      ws.cell(allCities.indexOf(country) + 3, 1)
        .string(country.name.common)
        .style(basicStyle);

      // Write Capital
      let textCapital = '';
      if(country.capital)
        textCapital = country.capital;
      else
        textCapital = '-';

      ws.cell(allCities.indexOf(country) + 3, 2)
      .string(textCapital)
      .style(basicStyle);
        
      // Write Area
      let textArea = '';
      if(country.area)
        textArea = country.area;
      else
        textArea = '-';
      
      ws.cell(allCities.indexOf(country) + 3, 3)
      .number(textArea)
      .style(basicStyle);

      // Write Currencies
      let textCurrencies = '';

      if(!country.currencies)
        textCurrencies = '-';
      else 
        textCurrencies = Object.keys(country.currencies).toString();

      ws.cell(allCities.indexOf(country) + 3, 4)
        .string(textCurrencies)
        .style(basicStyle);
    });
    
  } catch (error) {
    throw new Error(error.message || "Error trying to make request!");
  } 
  
  wb.write('countries_report.xlsx', res);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\nServer is running on port ${PORT}.`);
});
