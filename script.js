const fetch = require("node-fetch"); // Menggunakan node-fetch untuk request API
const xlsx = require("xlsx"); // Paket untuk manipulasi file Excel
const fs = require("fs");

// API key AccuWeather
const API_KEY = "e9cd3e6ce69697256eaafe8aceb8abfd";

// Function to fetch air quality data for a specific date
async function getAirQualityData(locationKey, date) {
  const API_URL = `http://dataservice.accuweather.com/airquality/v1/daily/1day/${locationKey}?apikey=${API_KEY}&details=true`;
  
  try {
    const response = await fetch(API_URL);
    const data = await response.json();

    // Logging data to check the structure
    console.log(`Data for ${date}:`, JSON.stringify(data, null, 2));

    // Ensure the data structure matches the expected format
    return {
      date: date,
      airQualityIndex: data[0]?.AirAndPollen?.[0]?.Value || "N/A", // Check if AirAndPollen exists
      category: data[0]?.AirAndPollen?.[0]?.Category || "N/A",
    };
  } catch (error) {
    console.error(`Error fetching data for ${date}:`, error);
    return null;
  }
}

// Function to generate date range from January to September
function generateDateRange() {
  const startDate = new Date(2024, 0, 1); // January 1, 2024
  const endDate = new Date(2024, 8, 30); // September 30, 2024
  const dates = [];

  for (let d = startDate; d <= endDate; d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d));
  }

  return dates;
}

// Main function to fetch data and write to Excel
async function generateAirQualityReport(locationKey) {
  const dates = generateDateRange();
  const airQualityData = [];

  for (let date of dates) {
    const formattedDate = date.toISOString().split("T")[0]; // Format date as YYYY-MM-DD
    const data = await getAirQualityData(locationKey, formattedDate);
    if (data) airQualityData.push(data);
  }

  // Convert data to worksheet and workbook
  const ws = xlsx.utils.json_to_sheet(airQualityData);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Air Quality Report");

  // Write to Excel file
  xlsx.writeFile(wb, "Air_Quality_Report_2024.xlsx");
  console.log("Data saved to Air_Quality_Report_2024.xlsx");
}

// Replace 'locationKey' with Surabaya's location key
const locationKey = "203449"; // Surabaya location key
generateAirQualityReport(locationKey);
