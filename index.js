const XlsxPopulate = require('xlsx-populate');
const express= require('express')
const app = express();
var data = require("./workForce");
var wageData = require("./wageData");

const port = 3030;

function rangeStyle(value, color, range, workbook, alignment = "center") {

    let row = workbook.sheet(0).range(range);
    workbook.sheet(0).column("A").width(10)
    workbook.sheet(0).column("B").width(35)
    workbook.sheet(0).column("C").width(35)
    workbook.sheet(0).column("D").width(12)
    workbook.sheet(0).column("E").width(15)
    workbook.sheet(0).column("F").width(12)
    workbook.sheet(0).column("G").width(12)
    workbook.sheet(0).column("H").width(35)
    row.value(value);
    row.style({ horizontalAlignment: alignment, verticalAlignment: "center" })
    row.style("fill", {
        type: "pattern",
        pattern: "darkDown",
        foreground: {
            rgb: color
        },
    });
    row.style("border", true);
    row.merged(true);
}

async function testMerge() {

    var workbook = await XlsxPopulate.fromBlankAsync();

    rangeStyle("Sr. No.", "61B33B", "A1:A3", workbook)
    rangeStyle("Name of the contractor", "61B33B", "B1:B3", workbook)
    rangeStyle("Description of work", "61B33B", "C1:C3", workbook)
    rangeStyle("TARGETS", "696969", "D1:G1", workbook)
    rangeStyle("CONSOLIDATED TOTAL NO.S", "61B33B", "H1:H3", workbook)
    rangeStyle(270, "696969", "D2:D2", workbook)
    rangeStyle(100, "", "E2:E2", workbook)
    rangeStyle(200, "696969", "F2:F2", workbook)
    rangeStyle(30, "696969", "G2:G2", workbook)
    rangeStyle("UN-SKILLED", "696969", "D3:D3", workbook)
    rangeStyle("SEMI-SKILLED", "696969", "E3:E3", workbook)
    rangeStyle("SKILLED", "696969", "F3:F3", workbook)
    rangeStyle("NONE", "696969", "G3:G3", workbook)

    data.forEach((row, i) => {
        rangeStyle(i+1, "ffffff", `A${i + 4}:A${i + 4}`, workbook, "right")
        rangeStyle(row.name, "ffffff", `B${i + 4}:B${i + 4}`, workbook, "left")
        rangeStyle(row.desc, "ffffff", `C${i + 4}:C${i + 4}`, workbook, "left")
        rangeStyle(row.unskilled, "ffffff", `D${i + 4}:D${i + 4}`, workbook)
        rangeStyle(row.semiskilled, "ffffff", `E${i + 4}:E${i + 4}`, workbook)
        rangeStyle(row.skilled, "ffffff", `F${i + 4}:F${i + 4}`, workbook)
        rangeStyle(row.none, "ffffff", `G${i + 4}:G${i + 4}`, workbook)
        rangeStyle(row.unskilled + row.semiskilled + row.skilled + row.none, "61B33B", `H${i + 4}:H${i + 4}`, workbook)
    })
    // Total
    rangeStyle("Total", "FFC0CB", `A${4 + data.length}:C${4 + data.length}`, workbook)
    rangeStyle(data.reduce((total, row) => total + row.unskilled, 0), "FED8B1", `D${4 + data.length}:D${4 + data.length}`, workbook)
    rangeStyle(data.reduce((total, row) => total + row.semiskilled, 0), "FED8B1", `E${4 + data.length}:E${4 + data.length}`, workbook)
    rangeStyle(data.reduce((total, row) => total + row.skilled, 0), "FED8B1", `F${4 + data.length}:F${4 + data.length}`, workbook)
    rangeStyle(data.reduce((total, row) => total + row.none, 0), "FED8B1", `G${4 + data.length}:G${4 + data.length}`, workbook)
    rangeStyle(data.reduce((total, row) => total + row.none + row.unskilled + row.skilled + row.semiskilled, 0), "61B33B", `H${4 + data.length}:H${4 + data.length}`, workbook)

    // Percent
    rangeStyle("% ACHIEVED", "FFC0CB", `A${5 + data.length}:C${5 + data.length}`, workbook)
    rangeStyle(((data.reduce((total, row) => total + row.unskilled, 0) / 270) * 100).toFixed(2)+"%", "FED8B1", `D${5 + data.length}:D${5 + data.length}`, workbook)
    rangeStyle(((data.reduce((total, row) => total + row.semiskilled, 0) / 100) * 100).toFixed(2)+"%", "FED8B1", `E${5 + data.length}:E${5 + data.length}`, workbook)
    rangeStyle(((data.reduce((total, row) => total + row.skilled, 0) / 200) * 100).toFixed(2)+"%", "FED8B1", `F${5 + data.length}:F${5 + data.length}`, workbook)
    rangeStyle(((data.reduce((total, row) => total + row.none, 0) / 30) * 100).toFixed(2)+"%", "FED8B1", `G${5 + data.length}:G${5 + data.length}`, workbook)

    workbook.toFileAsync(`./workforce.xlsx`);

}




//fetch all excel file data
app.get('/workforce', async(req, res) => {
    try {
    testMerge();
    setTimeout(()=>{
    res.download(`./workforce.xlsx`)
    },1000)
   
    } catch (error) {
        console.log(error);
    }
 })                


function rangleStyle2(value, color, range, workbook, isBold = false , setFont = 11 , alignment = "center") {

    let row = workbook.sheet(0).range(range);
    row.value(value);
    row.style({ horizontalAlignment: alignment, verticalAlignment: "bottom", fontSize: setFont})
    row.style("fill", {
        type: "pattern",
        pattern: "darkDown",
        foreground: {
            rgb: color
        },
    });
    row.style("border", true);

    if (isBold) {
        row.style("bold", true);
    }
    row.merged(true);
}


async function testMerge2() {

    var workbook = await XlsxPopulate.fromBlankAsync();

    rangleStyle2("FORM -- XVII", "FFFFFF", "A1:R1", workbook, true ,15)
    rangleStyle2("REGISTER OF WAGES", "FFFFFF", "A2:R2", workbook, true ,15)
    rangleStyle2("[(See Rule 78 (1) (a) (i)]of Contract Labour (Reg.& Abolition) Central & A.P.Rules)", "FFFFFF", "A3:R3", workbook)
    rangleStyle2("NAME AND ADDRESS OF THE CONTRACTOR:", "FFFFFF", "A4:B4", workbook, true)
    rangleStyle2("NAME AND ADDRESS OF THE ESTABLISHMENT IN/UNDER WHICH CONTRACT IS CARRIED ON:", "FFFFFF", "A5:B5", workbook, true)
    rangleStyle2("NATURE AND LOCATION OF WORK:", "FFFFFF", "A6:B6", workbook, true )
    rangleStyle2("NAME AND ADDRESS OF PRINCIPAL EMPLOYER:", "FFFFFF", "A7:B7", workbook, true)
    rangleStyle2("ACTION GUARDING SERVICES PVT LTD,HYD", "FFFFFF", "D4:S4", workbook, true ,"", "left")
    rangleStyle2("SECURITY,ALEKHYA - RISE,HYD", "FFFFFF", "D6:P6", workbook, true ,"", "left")
    rangleStyle2("WAGE PERIOD :", "FFFFFF", "Q6:Q6", workbook, true)
    rangleStyle2("MONTHLY", "FFFFFF", "R6:R6", workbook, true)
    rangleStyle2("9-2022", "FF0000", "S6:S6", workbook, true)
    rangleStyle2("SL N", "FFFFFF", "A9:A11", workbook, true)
    rangleStyle2("NAME OF THE WORKMAN", "FFFFFF", "B9:B11", workbook, true)
    rangleStyle2("EMPLOYEE ID", "FFFFFF", "C9:C11", workbook, true)
    rangleStyle2("DESIGNATION/NATURE OF WORK DONE", "FFFFFF", "D9:D11", workbook, true)
    rangleStyle2("NO OF DAYS", "FFFFFF", "E9:E11", workbook, true)
    rangleStyle2("UNIT OF WORK DONE", "FFFFFF", "F9:F11", workbook, true)
    rangleStyle2("DAILY RATE OF WAGES", "FFFFFF", "G9:G11", workbook, true)
    rangleStyle2("AMOUNT OF WAGES EARNED", "FFFFFF", "H9:L9", workbook, true)
    rangleStyle2("DEDUCTIONS IF ANY(INDICATE NATURE)P.F", "FFFFFF", "M9:P9", workbook, true)
    rangleStyle2("BASIC WAGE", "FFFFFF", "H10:H11", workbook, true)
    rangleStyle2("HRA", "FFFFFF", "I10:I11", workbook, true)
    rangleStyle2("SITE ALLOUNCE", "FFFFFF", "J10:J11", workbook, true)
    rangleStyle2("OTHER CASH PAYMENT", "FFFFFF", "K10:K11", workbook, true)
    rangleStyle2("TOTAL", "FFFFFF", "L10:L11", workbook, true)
    rangleStyle2("ESI", "FFFFFF", "M10:M11", workbook, true)
    rangleStyle2("PF", "FFFFFF", "N10:N11", workbook, true)
    rangleStyle2("PT", "FFFFFF", "O10:O11", workbook, true)
    rangleStyle2("TOTAL DEDUCTIONS", "FFFFFF", "P10:P11", workbook, true)
    rangleStyle2("NET AMOUNT PAID", "FFFFFF", "Q9:Q11", workbook, true)
    rangleStyle2("SIGNATURE", "FFFFFF", "R9:R11", workbook, true)
    rangleStyle2("INITIAL OF CONTRACTOR", "FFFFFF", "S9:S11", workbook, true)

    wageData.map((ele, i) => {
        rangleStyle2(i+1, "FFFFFF", `A${12 + i}:A${12 + i}`, workbook)
        rangleStyle2(ele.name, "FFFFFF", `B${12 + i}:B${12 + i}`, workbook)
        rangleStyle2(ele.empid, "FFFFFF", `C${12 + i}:C${12 + i}`, workbook)
        rangleStyle2(ele.des, "FFFFFF", `D${12 + i}:D${12 + i}`, workbook)
        rangleStyle2(ele.day, "FFFFFF", `E${12 + i}:E${12 + i}`, workbook)
        rangleStyle2(ele.unit, "FFFFFF", `F${12 + i}:F${12 + i}`, workbook)
        rangleStyle2(ele.daily, "FFFFFF", `G${12 + i}:G${12 + i}`, workbook)
        rangleStyle2(ele.basic, "FFFFFF", `H${12 + i}:H${12 + i}`, workbook)
        rangleStyle2(ele.hra, "FFFFFF", `I${12 + i}:I${12 + i}`, workbook)
        rangleStyle2(ele.site, "FFFFFF", `J${12 + i}:J${12 + i}`, workbook)
        rangleStyle2(ele.other, "FFFFFF", `K${12 + i}:K${12 + i}`, workbook)
        rangleStyle2(ele.total, "FFFFFF", `L${12 + i}:L${12 + i}`, workbook)
        rangleStyle2(ele.esi, "FFFFFF", `M${12 + i}:M${12 + i}`, workbook)
        rangleStyle2(ele.pf, "FFFFFF", `N${12 + i}:N${12 + i}`, workbook)
        rangleStyle2(ele.pt, "FFFFFF", `O${12 + i}:O${12 + i}`, workbook)
        rangleStyle2(ele.td, "FFFFFF", `P${12 + i}:P${12 + i}`, workbook)
        rangleStyle2(ele.net, "FFFFFF", `Q${12 + i}:Q${12 + i}`, workbook)
        rangleStyle2(ele.sign, "FFFFFF", `R${12 + i}:R${12 + i}`, workbook)
        rangleStyle2(ele.initial, "FFFFFF", `S${12 + i}:S${12 + i}`, workbook)
    })

    workbook.toFileAsync(`./wages.xlsx`);
}


//fetch all excel file data
app.get('/wages', async(req, res) => {
    try {
    testMerge2();
    setTimeout(()=>{
        res.download(`./wages.xlsx`)
    },2000)
   
    } catch (error) {
        console.log(error);
    }
 })      

 
app.listen(port, () => {
    console.log(`Server listening on port ${port}`);
});
