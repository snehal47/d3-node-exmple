let express = require('express');
var path = require('path');
const canvas = require('canvas');
const D3Node = require('d3-node');
const docx = require("docx");
const fs = require("fs");
const { Document, Packer, Paragraph, ImageRun, HeadingLevel, SectionType, TableOfContents }  = docx;

var app = express();

app.set('view engine', 'html');
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));

async function getImage (d3canvas, svgDataUrl) {
  return new Promise(function (resolve, reject) {
    const canvasImage = new canvas.Image();
    canvasImage.onload = function () {
      d3canvas.getContext('2d').drawImage(canvasImage, 0, 0);
      resolve("OK");
    };
    canvasImage.onerror = function () {
      reject("Error while creating Image");
    };
    canvasImage.src = svgDataUrl;
  });
};

app.get('/', (req, res, next) => {
  (async () => {
    const d3n = new D3Node({ canvasModule: canvas });
    const d3canvas = d3n.createCanvas(200, 200);
    const svg = d3n.createSVG(200, 200);
    svg
      .selectAll('circles')
      .data([
        { type: 'small', r: 25, x: 25, y: 100, color: '#606060' },
        { type: 'medium', r: 50, x: 75, y: 100, color: '#378B91' },
        { type: 'large', r: 75, x: 125, y: 100, color: '#E5CE4E' }
      ])
      .enter()
      .append('circle')
      .attr('r', (d) => d.r)
      .attr('cx', (d) => d.x)
      .attr('cy', (d) => d.y)
      .attr('fill', (d) => d.color)
      .attr('class', (d) => d.type);
    const status = await getImage(d3canvas, 'data:image/svg+xml;charset=utf8,' + d3n.svgString());
    if(status === "OK"){
      const img = Buffer.from(d3canvas.toDataURL('image/png').split(',')[1], 'base64');
      const doc = new Document({
        features: { updateFields: true },
        sections: [
            {
                properties: { titlePage: true, },
                children: [new Paragraph({ text: "XXX Report 2022", heading: HeadingLevel.TITLE }),]
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ text: "XXX Report 2022", heading: HeadingLevel.HEADING_1, }), new Paragraph({ text: "[Provide a description of your company and your sustainability policy here.]", }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new TableOfContents("Table of contents", { hyperlink: true, headingStyleRange: "1-5", }),]
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [
                    new Paragraph({ heading: HeadingLevel.HEADING_1, text: "XXX 102", }),
                    new Paragraph({ text: "DISCLOSURE 102-1", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({children: [new ImageRun({data: img, transformation: {width: [200],height: [200]}})]}),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102-3", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102-4", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102-5", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102-6", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                    new Paragraph({ text: "DISCLOSURE 102-7", heading: HeadingLevel.HEADING_5, }),
                    new Paragraph({ text: "[Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.] ", }),
                ],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "1. PROFILE " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "2. STRATEGY " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "3. INTEXXXTY " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "4. GOVERNANCE " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "5. ENGAGEMENT" }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_4, text: "6. PRACTICE " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_1, text: "XXX 200: ECONOMIC " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 201: PERFORMANCE" }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 205: CORRUPTION " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 206: BEHAVIOUR " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_1, text: "XXX 300: ENVIRONMENTAL " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 301: MATERIALS " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 302: ENERGY " }),],
            },
            {
                properties: { type: SectionType.NEXT_PAGE, },
                children: [new Paragraph({ heading: HeadingLevel.HEADING_3, text: "XXX 303: WATER " }),],
            }
        ]
      });
      
      Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("MyDocument.docx", buffer);
      });

      return res.send({ error: false, file: `http://localhost:8000/MyDocument.docx`});
    }else{
      return res.send({ error: true, message: status});
    }
    
  })().catch(next);
});

let listener = app.listen(8080, function () {
  console.log('Listening on port ' + listener.address().port);
});
