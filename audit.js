const fs = require('fs');

const appJs = fs.readFileSync('c:/Users/nnkre/keisokuGoogle/sidepanel/app.js', 'utf8');
const indexHtml = fs.readFileSync('c:/Users/nnkre/keisokuGoogle/sidepanel/index.html', 'utf8');

const idRegex = /document\.getElementById\(['"]([^'"]+)['"]\)/g;
let match;
const appIds = new Set();
while ((match = idRegex.exec(appJs)) !== null) {
    appIds.add(match[1]);
}

const htmlIdRegex = /id=["']([^"']+)["']/g;
const htmlIds = new Set();
while ((match = htmlIdRegex.exec(indexHtml)) !== null) {
    htmlIds.add(match[1]);
}

const missingInHtml = [...appIds].filter(id => !htmlIds.has(id));
console.log('IDs used in app.js but missing in index.html:', missingInHtml);

const wizardJs = fs.readFileSync('c:/Users/nnkre/keisokuGoogle/sidepanel/wizard.js', 'utf8');
const appJsImports = appJs.match(/import\s+\{[^\}]+\}\s+from\s+[^;]+;/g) || [];
const wizardJsImports = wizardJs.match(/import\s+\{[^\}]+\}\s+from\s+[^;]+;/g) || [];

console.log('\n--- Duplicate Imports in app.js ---');
const appImportMap = {};
appJsImports.forEach(imp => {
    const matches = imp.match(/\{([^\}]+)\}/);
    if (matches) {
        matches[1].split(',').forEach(i => {
            const name = i.trim();
            if (appImportMap[name]) console.log('Duplicate:', name);
            appImportMap[name] = true;
        });
    }
});

console.log('\n--- Duplicate Imports in wizard.js ---');
const wizImportMap = {};
wizardJsImports.forEach(imp => {
    const matches = imp.match(/\{([^\}]+)\}/);
    if (matches) {
        matches[1].split(',').forEach(i => {
            const name = i.trim();
            if (wizImportMap[name]) console.log('Duplicate:', name);
            wizImportMap[name] = true;
        });
    }
});
