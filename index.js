function calculateCOG(lat1, lon1, lat2, lon2) {
  const deltaLon = lon2 - lon1;
  const deltaLat = lat2 - lat1;

  // Calculate COG in degrees
  const cog = (Math.atan2(deltaLon, deltaLat) * (180 / Math.PI)) % 360;

  // Normalize the COG to be between 0 and 360 degrees
  return (cog + 360) % 360;
}

const speed_ms_to_knots = (speed) => speed * 1.94384;

const firebaseConfig = {
  apiKey: "AIzaSyCr3V4xwWEP0yTdYRcvt4vzuKNdUAEB5bI",

  authDomain: "kapallawd-40b01.firebaseapp.com",

  projectId: "kapallawd-40b01",

  storageBucket: "kapallawd-40b01.appspot.com",

  messagingSenderId: "383307595328",

  appId: "1:383307595328:web:e3cf259d7c707e9de629a5",
};

import { initializeApp } from "firebase/app";
import {
  getFirestore,
  collection,
  getDocs,
  orderBy,
  query,
} from "firebase/firestore";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const gpsRef = collection(db, "gps");

function formatTimestamp(timestamp) {
  const date = timestamp.toDate();
  date.setHours(date.getHours() + 7);
  return date.toISOString().replace("T", " ").split(".")[0]; // Format as "YYYY-MM-DD HH:MM:SS"
}

async function exportToExcel() {
  try {
    const q = query(gpsRef, orderBy("timestamp", "asc"));
    const snapshot = await getDocs(q);

    if (snapshot.empty) {
      console.log("No data found");
      return;
    }

    // Create a new Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Firestore Data");

    // Add header row based on Firestore document structure
    worksheet.columns = [
      { header: "Timestamp", key: "timestamp", width: 20 },
      { header: "Latitude", key: "lat", width: 15 },
      { header: "Longitude", key: "long", width: 15 },
      { header: "Speed (m/s)", key: "speed_ms", width: 15 },
      { header: "Course Over Ground (COG)", key: "cog", width: 20 },
      { header: "Speed (knots)", key: "speed_knots", width: 15 },
    ];

    // Add data rows from Firestore documents
    let prevData = null;
    snapshot.forEach((doc) => {
      const data = doc.data();
      const speed_knots = speed_ms_to_knots(data.speed);
      const readableTimestamp = formatTimestamp(data.timestamp);

      let cog = null;
      if (prevData) {
        cog = calculateCOG(
          prevData.loc._lat,
          prevData.loc._long,
          data.loc._lat,
          data.loc._long
        );
      }

      worksheet.addRow({
        timestamp: readableTimestamp,
        lat: data.loc._lat,
        long: data.loc._long,
        speed_ms: data.speed,
        cog: cog,
        speed_knots: speed_knots,
      });

      prevData = data;
    });

    // Save the workbook to a file
    const filePath = path.join("data.xlsx");
    await workbook.xlsx.writeFile(filePath);

    console.log(`Excel file has been created at ${filePath}`);
  } catch (error) {
    console.error("Error exporting to Excel:", error);
  }
}

// Run the function
exportToExcel();
