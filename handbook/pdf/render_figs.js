const fs = require("fs");
const path = require("path");
const vm = require("vm");
const root = path.resolve(__dirname, "..");
const fig = path.join(__dirname, "fig");
fs.mkdirSync(fig, { recursive: true });
const ctx = { console, Math };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, "js/data.js"), "utf8"), ctx);
vm.runInContext(fs.readFileSync(path.join(root, "js/draw.js"), "utf8"), ctx);
function dump(file, expr) {
  fs.writeFileSync(path.join(fig, file), vm.runInContext(expr, ctx));
}
dump("narrow.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b737"), {mode:"detail", ppm:68})');
dump("b767.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b767"), {mode:"detail", ppm:58})');
dump("b787.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b787"), {mode:"detail", ppm:56})');
dump("b747.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b747"), {mode:"detail", ppm:48})');
dump("narrow-mini.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b737"), {mode:"gallery", ppm:52})');
dump("b767-mini.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b767"), {mode:"gallery", ppm:40})');
dump("b787-mini.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b787"), {mode:"gallery", ppm:38})');
dump("b747-mini.svg", 'crossSectionSVG(AIRCRAFT.find(a=>a.id==="b747"), {mode:"gallery", ppm:32})');
dump("gear737.svg", 'gearSVG(AIRCRAFT.find(a=>a.id==="b737"))');
dump("gear757.svg", 'gearSVG(AIRCRAFT.find(a=>a.id==="b757"))');
dump("gear747.svg", 'gearSVG(AIRCRAFT.find(a=>a.id==="b747"))');
dump("gear767.svg", 'gearSVG(AIRCRAFT.find(a=>a.id==="b767"))');
dump("gear787.svg", 'gearSVG(AIRCRAFT.find(a=>a.id==="b787"))');
