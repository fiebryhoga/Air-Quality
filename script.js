const fetch = require("node-fetch");
const xlsx = require("xlsx"); 
const fs = require("fs");

// API key AccuWeather
const API_KEY = "e9cd3e6ce69697256eaafe8aceb8abfd";

async function getAirQualityData(locationKey, date) {
  const API_URL = `http://dataservice.accuweather.com/airquality/v1/daily/1day/${locationKey}?apikey=${API_KEY}&details=true`;
  
  try {
    const response = await fetch(API_URL);
    const data = await response.json();

    console.log(`Data for ${date}:`, JSON.stringify(data, null, 2));

    return {
      date: date,
      airQualityIndex: data[0]?.AirAndPollen?.[0]?.Value || "N/A", 
      category: data[0]?.AirAndPollen?.[0]?.Category || "N/A",
    };
  } catch (error) {
    console.error(`Error fetching data for ${date}:`, error);
    return null;
  }
}

function generateDateRange() {
  const startDate = new Date(2024, 0, 1); 
  const endDate = new Date(2024, 8, 30); 
  const dates = [];

  for (let d = startDate; d <= endDate; d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d));
  }

  return dates;
}

//function buat ngubah ke excel
async function generateAirQualityReport(locationKey) {
  const dates = generateDateRange();
  const airQualityData = [];

  for (let date of dates) {
    const formattedDate = date.toISOString().split("T")[0];
    const data = await getAirQualityData(locationKey, formattedDate);
    if (data) airQualityData.push(data);
  }

  const ws = xlsx.utils.json_to_sheet(airQualityData);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Air Quality Report");

  xlsx.writeFile(wb, "Air_Quality_Report_2024.xlsx");
  console.log("Data saved to Air_Quality_Report_2024.xlsx");
}

const locationKey = "203449"; // Surabaya location key
generateAirQualityReport(locationKey);
