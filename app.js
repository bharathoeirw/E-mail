const express = require("express");
const Sequelize = require("sequelize");
const connection = require("./Database/connection");
const Employee = require("./Model/EmployeeDetails");
const XLSX = require("xlsx");
const fs = require("fs");
const archiver = require("archiver");
const path = require("path");
const nodemailer = require("nodemailer");

const port = 5555;
const app = express();
const router = express.Router();

app.use(express.json());

app.use(router);

router.post("/exportData", async (req, res) => {
  const { from, to } = req.body;

  if (!from || !to) {
    return res.status(400).send("Both 'from' and 'to' email addresses are required.");
  }

  try {
    const employees = await Employee.findAll();
    const dataRows = employees.map((row) => [
      row.id,
      row.firstName,
      row.lastName,
      row.email,
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([
      ["ID", "First Name", "Last Name", "Email"],
      ...dataRows,
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Employee Data");

    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    const zipFileName = "EmployeeReport.zip";
    const zipFilePath = path.resolve(__dirname, zipFileName);

    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver("zip", {
      zlib: { level: 9 },
    });

    output.on("close", function () {
      console.log(archive.pointer() + " total bytes");
      console.log("File was successfully compressed");

      // Send email with the zip file attachment
      sendEmail(from, to, zipFilePath)
        .then(() => {
          res.send("Email sent successfully");
        })
        .catch((error) => {
          console.error("Error sending email:", error);
          res.status(500).send("Failed to send email");
        });
    });

    archive.on("warning", function (err) {
      if (err.code === "ENOENT") {
        console.warn(err);
      } else {
        throw err;
      }
    });

    archive.on("error", function (err) {
      throw err;
    });

    archive.pipe(output);
    archive.append(buffer, { name: "EmployeeData.xlsx" });

    archive.finalize();
  } catch (error) {
    console.error("Error exporting data:", error);
    res.status(500).send("Internal Server Error");
  }
});

connection
  .sync()
  .then((result) => {
    console.log("Database connected successfully");
    app.listen(port, () => {
      console.log(`Server is running ${port}`);
    });
  })
  .catch((error) => {
    console.log("Unable to connect Database");
  });

async function sendEmail(from, to, filePath) {

  let transporter = nodemailer.createTransport({
    host: "smtp.gmail.com", 
    port: 587,
    secure: false, 
    auth: {
      user: "bharathkumar1027@gmail.com", 
      pass: "fijx vsuu cpkb eqdi", 
    },
  });
  


  // send mail with defined transport object
  let info = await transporter.sendMail({
    from: from,
    to: to,
    subject: "Employee Report",
    text: "Please find the attached employee report.",
    attachments: [
      {
        filename: "EmployeeReport.zip",
        path: filePath,
      },
    ],
  });

  console.log("Message sent: %s", info.messageId);
}
