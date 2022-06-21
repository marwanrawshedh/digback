const express = require('express');
const app = express();
const cors = require('cors');
const multer = require('multer');

const fsExtra = require('fs-extra')
const { readSheetNames } = require('read-excel-file/node')
const readXlsxFile = require('read-excel-file/node');

app.use(cors())
app.use(express.json());

let fileName
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, './uploads')
  },
  filename: function (req, file, cb) {

    cb(null, fileName + '.xlsx')
  }
})
const upload = multer({ storage: storage }).single('file')

app.post('/', (req, res) => {
  fileName = req.query.name


  upload(req, res, function (err) {
    if (err instanceof multer.MulterError) {
      // A Multer error occurred when uploading.
    } else if (err) {
      // An unknown error occurred when uploading.
    }

    readSheetNames(`${__dirname}/uploads/${fileName}.xlsx`).then((sheetNames) => {

      readXlsxFile(`${__dirname}/uploads/${fileName}.xlsx`).then((row) => {
        // `rows` is an array of rows
        // each row being an array of cells.
        const rowL = new Array(row.length);
        res.status(200).send({ sheetNames, rowL })
      })
    })
  })



})
// app.get("/delete", () => {
//   console.log("deleting");
//   fsExtra.emptyDirSync(`${__dirname}/uploads`)


// })
app.post('/compare', async (req, res) => {
  let { sheet1, sheet2, name1, name2, firstempl, secempl, firstcompare, seccompare, row1, row2,numHours } = req.body
  name1 = name1.toLowerCase().includes("leave") ? "Leaves File" : "Time Sheet File"
  name2 = name2.toLowerCase().includes("leave") ? "Leaves File" : "Time Sheet File"



  let fileOne = {};
  const fileTwo = []
  let dif1 = []
  let notEqual = []
  let final1 = []
  let final2 = []


  let sheetFile = await readXlsxFile(`${__dirname}/uploads/first.xlsx`, { sheet: `${sheet1}` })



  let sheetFile2 = await readXlsxFile(`${__dirname}/uploads/sec.xlsx`, { sheet: `${sheet2}` })


  for (let i = row1; i < sheetFile.length; i++) {

    fileOne[convert(sheetFile[i][firstempl])] = [i + 1, ...sheetFile[i]]


  };

  for (let i = row2; i < sheetFile2.length; i++) {

    sheetFile2[i].unshift(i + 1)

    if (!fileOne[convert(sheetFile2[i][secempl + 1])]) {

      fileTwo.push(sheetFile2[i])
    } else {
      if (fileOne[convert(sheetFile2[i][secempl + 1])][firstcompare + 1] != sheetFile2[i][seccompare + 1]) {
        let arr1 = [0, firstempl + 1, firstcompare + 1

        ]
        let arr2 = [0, secempl + 1, seccompare + 1,

        ]


        dif1.push([...format(arr1, fileOne[convert(sheetFile2[i][secempl + 1])]), ...format(arr2, sheetFile2[i]).reverse()])
      } else if (fileOne[convert(sheetFile2[i][secempl + 1])][firstcompare + 1] === sheetFile2[i][seccompare + 1] && sheetFile2[i][seccompare + 1] != numHours) {
        let arr1 = [0, firstempl + 1, firstcompare + 1

        ]
        let arr2 = [0, secempl + 1, seccompare + 1,

        ]


        notEqual.push([...format(arr1, fileOne[convert(sheetFile2[i][secempl + 1])]), ...format(arr2, sheetFile2[i]).reverse()])
      }


    }

    delete fileOne[convert(sheetFile2[i][secempl + 1])]

  }


  fsExtra.emptyDirSync(`${__dirname}/uploads`)
  fileOne = Object.values(fileOne)



  final1.unshift([`Row Number in ${name1}`, 'Employee Name'])
  final2.unshift([`Row Number in ${name2}`, 'Employee Name'])





  fileOne.forEach((element) => {
    final1.push([...format([0, firstempl + 1], element)])
  })
  fileTwo.forEach((element) => {
    final2.push([...format([0, secempl + 1], element)])
  })


  dif1.unshift([`Row Number in ${name1}`, 'Employee Name', `Total Hour in ${name1}`, `Total Hour in ${name2}`, 'Employee Name', `Row Number in ${name2}`])
  notEqual.unshift([`Row Number in ${name1}`, 'Employee Name', `Total Hour in ${name1}`, `Total Hour in ${name2}`, 'Employee Name', `Row Number in ${name2}`])
  res.status(200).send({ fileOne: final1, fileTwo: final2, fileDif: dif1, notEqual })

})




app.listen(process.env.PORT||3010, () => {
  console.log('Server started on port 3010');
})


function format(arr, data) {

  let res = []
  data.forEach((element, i) => {
    if (arr.includes(i)) {
      res.push(element)
    }

  })
  return res
}
let convert = (str) => {
  try {

    return str.replace(/[^\w]/g, "").toLowerCase()
  } catch (error) {
    return str

  }


}

