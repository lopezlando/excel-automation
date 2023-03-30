//libraries
const xlsx = require("xlsx");
const { format, parse, isValid, parseISO } = require("date-fns");

console.log("Reading file...");

// retrieves the desired file
const book1 = xlsx.readFile(
  "PDI_Daily Report SSAS V1 (version 2)(Recuperado automáticamente).xlsm"
);

console.log("Done");

//creates a json object based on the excel sheet
const data = xlsx.utils.sheet_to_json(book1.Sheets["MaestroFull"], {
  range: 1,
  defval: "",
});

//columns to delete from the json object
const keysToDelete = [
  "Leche medicamentosa",
  "Columna1",
  "total emitido",
  "Total CD",
  "__EMPTY",
  "__EMPTY_1",
  "__EMPTY_2",
  "__EMPTY_3",
];

// these columns must be re-converted to dates, JavaScript makes them numbers
const numbersToDates = [
  "ULTODC170",
  "ULTODC171",
  "ULTODC179",
  "ULTODC262",
  "ULTODC311",
  "Ultimo ingresoCABA",
  "Ultimo ingresoCBA",
  "FECHAALTA",
];

console.log("Deleting irrelevant columns...");

//deletion of the columns and conversion of dates
data.forEach((row) => {
  keysToDelete.forEach((key) => delete row[key]);
  row["Fecha daily"] = format(new Date(Date.now()), "dd/MM/yyyy");

  // converts dates (JavaScript auto converts them to numbers, so we have to revert that)
  Object.keys(row).forEach((key) => {
    if (numbersToDates.includes(key)) {
      const dateValue = row[key];

      // Check whether it is a valid date value or not
      if (typeof dateValue !== "number") {
        row[key] = "";
      } else if (!isNaN(dateValue)) {
        const dateMs = (dateValue - (25567 + 2)) * 86400 * 1000;
        const adjustedDateMs = dateMs + 86400000; // add one day in milliseconds
        const dateStr = new Date(adjustedDateMs).toISOString();
        row[key] = parseISO(dateStr);
      }
    }
  });
});

//the last object of the array is empty. pop() to delete it
data.pop();

console.log("Done");

console.log("Creating backup file...");

//creates a new worksheet using the updated data
const newWorksheet = xlsx.utils.json_to_sheet(data);

//new excel workbook
const newWorkbook = xlsx.utils.book_new();

//appends our worksheet to the new workbook
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Hoja1");

//creates a date for the file's title
const today = format(new Date(Date.now()), "dd-MM-yyyy");

//writes the new file
xlsx.writeFile(newWorkbook, `Seguimiento Daily/${today} Daily.xlsx`);

console.log("Done");
