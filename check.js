const fs = require('fs');
const html = fs.readFileSync('sidepanel/index.html', 'utf8');
const js = fs.readFileSync('sidepanel/app.js', 'utf8');
const matches = [...js.matchAll(/getElementById\('([^']+)'\)/g)];
const ids = matches.map(m => m[1]);
const missing = ids.filter(id => !html.includes('id="' + id + '"'));
console.log('Missing IDs:', missing);
