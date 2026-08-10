// +^j to open console
// doing it properly first

// locate Deposit page records matching the overhead and date, then mark them as Checked
function checkRecords(overhead, date) {
    let recordList = Array.prototype.slice.call(document.getElementsByClassName("uir-list-row-tr"))
    let relevantList = recordList.filter((x) => x.childNodes[4].innerHTML.includes(overhead) && x.childNodes[4].innerHTML.includes(date))
    relevantList.forEach((x) => x.childNodes[0].childNodes[0].click()) // element has listeners, so just setting .checked doesn't work
}

// make range of date strings from start/end date strings
function makeDateRange(start, end) {
    let outArr = []
    for(let d = new Date(start); d <= new Date(end); d.setDate(d.getDate() + 1)) {
        outArr.push((new Date(d)).toLocaleDateString(undefined, {month: "2-digit", day: "2-digit"}))
    }
    return outArr
}

function markRecords(overhead, dateList) {
    const startHigh = ["BLV", "09801", "01904", "PAS"]
    dateList.slice(0,-1).forEach((date) => checkRecords(overhead, date))
    if(startHigh.includes(overhead)) {
        checkRecords(overhead, dateList.slice(-1))
    }
}

// and here's the functionless way derived from those to put in an .ahk

// three placeholder vars
let overhead = "BLV";
let startDate = "08.02";
let endDate = "08.03";
const startHigh = ["BLV", "09801", "01904", "PAS"];
let dateList = [];
for(let d = new Date(startDate); d <= new Date(endDate); d.setDate(d.getDate() + 1)) {
    dateList.push(d.toLocaleDateString(undefined, {month: "2-digit", day: "2-digit"}));
}
if(! startHigh.includes(overhead)) {
    dateList = dateList.slice(0, -1);
}
for(const date of dateList) {
    let recordList = Array.prototype.slice.call(document.getElementsByClassName("uir-list-row-tr"));
    let relevantList = recordList.filter((x) => x.childNodes[4].innerHTML.includes(overhead) && x.childNodes[4].innerHTML.includes(date));
    relevantList.forEach((x) => x.childNodes[0].childNodes[0].click());
}
