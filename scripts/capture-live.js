// Excel live-control function body. Supply "const mapping = <resolved map JSON>;" above this file.
// Read after Excel has recalculated. The distributed master uses the 1900 date system.
const sheets = ctx.workbook.worksheets;
sheets.load("items/name");
const fields = [];
const anchors = [];
for (const [name, spec] of Object.entries(mapping)) {
  const sheet = sheets.getItem(name);
  for (const [address, label] of Object.entries(spec.anchors)) {
    const range = sheet.getRange(address);
    range.load("text");
    anchors.push({ name, address, label, range });
  }
  for (const [key, address] of Object.entries(spec.outputs)) {
    const range = sheet.getRange(address);
    range.load("values");
    fields.push({ key, range });
  }
}
await ctx.sync();
const normalize = value => String(value).trim().toLowerCase().replace(/\s+/g, " ");
for (const { name, address, label, range } of anchors) {
  const allowed = Array.isArray(label) ? label : [label];
  if (!allowed.map(normalize).includes(normalize(range.text[0][0]))) {
    throw Error("Resolve workbook map at " + name + "!" + address + " (" + label + ")");
  }
}
const result = {};
for (const { key, range } of fields) {
  let values = range.values.flat();
  if (["dates", "close", "sale", "history_end"].includes(key)) {
    values = values.map(v => typeof v === "number" ?
      new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10) : v);
  }
  result[key] = values.length === 1 ? values[0] : values;
}
const errors = sheets.items.map(sheet => {
  const ranges = sheet.getUsedRange().getSpecialCellsOrNullObject(
    Excel.SpecialCellType.formulas, Excel.SpecialCellValueType.errors);
  ranges.load("address");
  return ranges;
});
await ctx.sync();
result.formula_errors = errors.filter(r => !r.isNullObject).map(r => r.address);
return result;
