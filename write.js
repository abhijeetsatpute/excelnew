// // Requiring module
// const reader = require('xlsx')
  
// // Reading our test file
// const file = reader.readFile('./test.xlsx')
  
// // Sample data set
// let student_data = [{
//     Name:'Nikhil',
//     Age:22,
//     Branch:'ISE',
//     Marks: 70
// },
// {
//     Name:'Amitha',
//     Age:21,
//     Branch:'EC',
//     Marks:80
// }]
  
// const ws = reader.utils.json_to_sheet(student_data)
  
// reader.utils.book_append_sheet(file,ws,"Sheet3")
  
// // Writing to our file
// reader.writeFile(file,'./test.xlsx')

const reader = require('xlsx')
var file = reader.readFile('./include/dic.xlsx');
  
let data = {}
  
var sheets = file.SheetNames

for(let i = 0; i < sheets.length; i++)
{
   var temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
      data[res['Abbr.']] = res['Full']
   })
}

let writedata = []
file = reader.readFile('./include/map.xlsx');
sheets = file.SheetNames
for(let i = 0; i < sheets.length; i++)
{
   var temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
    res['Full form'] = ''
    var tempaddr = res['Abbr'].split(', ')
    tempaddr.forEach((eachaddr,index) => {
        if(index+1 < tempaddr.length){
            res['Full form'] += data[eachaddr]+', '
        }else{
            res['Full form'] += data[eachaddr]
        }
    })
    writedata.push(res)
   })
}

// Printing data

const ws = reader.utils.json_to_sheet(writedata)
  
reader.utils.book_append_sheet(file,ws,"Sheet4")
  
// Writing to our file
reader.writeFile(file,'./include/final.xlsx')