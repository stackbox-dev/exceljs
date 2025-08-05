/* eslint-disable */
const Excel = require('./lib/exceljs.nodejs');
const fs = require('fs');

async function* generateData() {
  // Generate 1 million rows
  const str = Math.random().toString(36).substring(7);
  const cols = Array.from({length: 50}, () => str);
  for (let i = 0; i < 10000; i++) {
    if (i % 1000 === 0) {
      await new Promise(resolve => setImmediate(resolve));
    }
    if (i % 100000 === 0) {
      console.log(`Generated ${i} rows`);
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
    yield cols;
  }
}

async function test1() {
  const workbook = new Excel.stream.xlsx.WorkbookWriter({
    filename: './huge.xlsx',
    useStyles: true,
    useSharedStrings: false,
    dontUseFsCapacitor: true,
  });

  const sheetProms = [];
  for (let i = 0; i < 5; i++) {
    sheetProms.push(addSheet(workbook, i));
  }

  const interval = setInterval(() => {
    global.gc();
    const memory = process.memoryUsage();
    console.log(`Memory usage: ${Math.round(memory.heapUsed / 1024 / 1024)}MB`);
  }, 1000);

  await Promise.all(sheetProms);
  await workbook.commit();

  clearInterval(interval);

  global.gc();
  const used = process.memoryUsage();
  console.log(`Memory usage: ${Math.round(used.heapUsed / 1024 / 1024)}MB`);
  console.log('Done!');

  fs.stat('./huge.xlsx', (err, stats) => {
    if (err) throw err;
    console.log(`File size: ${Math.round(stats.size / 1024 / 1024)}MB`);
  });
}
const startTime = Date.now();
test1().finally(() => {
  global.gc();
  const used = process.memoryUsage();
  console.log(`Memory usage: ${Math.round(used.heapUsed / 1024 / 1024)}MB`);
  const endTime = Date.now();
  console.log(`Total time: ${(endTime - startTime) / 1000} seconds`);
});
async function addSheet(workbook, i) {
  const worksheet = workbook.addWorksheet('Sheet' + (i + 1));

  // Process data using async generator
  for await (const data of generateData()) {
    worksheet.addRow(data).commit();
  }

  // Commit the worksheet and workbook
  await worksheet.commit();
}
