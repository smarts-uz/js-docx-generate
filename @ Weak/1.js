const fs = require("fs");
const {
  patchDocument,
  PatchType,
  TextRun,
} = require("docx");

const templateData = fs.readFileSync("template.docx");
const outputType = "nodebuffer";
const patches = {
  greeting: {
    type: PatchType.PARAGRAPH,
    children: [ new TextRun("Hello Dilbbekkkkk") ],
  },
};

patchDocument(templateData, { outputType, patches })
  .then((docBuffer) => {
    fs.writeFileSync("output1.docx", docBuffer);
    console.log("✅ Document patched and saved!");
  })
  .catch(err => console.error(err));
