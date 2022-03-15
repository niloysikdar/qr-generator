const qr = require("qr-image");
const readXlsxFile = require("read-excel-file/node");

const fileName = "data.xlsx";
const sheetName = "Form Responses 1";
const outDir = "out";

readXlsxFile(fileName, { sheet: sheetName }).then((rows) => {
  rows.shift();

  rows.map((studentData, index) => {
    const allData = {
      fullName: studentData[1],
      department: studentData[2],
      year: studentData[4],
      rollNo: studentData[3],
      dob: studentData[6],
      bloodGroup: studentData[5],
      contactNo: studentData[7],
    };

    const qrString = `Name: ${allData.fullName}, Department: ${allData.department}, Year: ${allData.year}, Roll No: ${allData.rollNo}, DOB: ${allData.dob}, Blood Group: ${allData.bloodGroup}, Contact No: ${allData.contactNo}`;

    console.log(`${index + 1}: Generating QR for ${allData.fullName}`);

    var qr_file = qr.image(qrString, { type: "png" });
    qr_file.pipe(
      require("fs").createWriteStream(`${outDir}/${allData.rollNo}.png`)
    );
  });
});
