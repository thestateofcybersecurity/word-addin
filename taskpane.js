/* global Office, Word, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insertTitlePageButton").onclick = insertTitlePage;
    document.getElementById("applyStylesButton").onclick = applyCustomStyles;
    document.getElementById("insertHeaderFooterButton").onclick = insertHeaderFooter;
    document.getElementById("generateTOCButton").onclick = insertTableOfContents;
    document.getElementById("importExcelTable").onclick = insertExcelTable;
    document.getElementById("addBiosButton").onclick = addEngagementTeamBios;
  }
});

// Employee bios data loaded from attached files
const bios = {
  "Alvin": `Alvin Tugume Cybersecurity Consultant: vCISO Infosec. Engineer Incident Responder Risk Assessor CCSKv4 | CompTIA SEC+ | TPN Certified

With 10 years in IT and Cybersecurity Alvin Tugume is a recognized expert in the field. A proud holder of a Bachelor's degree in Cybersecurity. Prior to Richey May Alvin was responsible for the cybersecurity posture of three credit unions and a call center. Responsibilities included maintaining compliance with information security policies, monitoring the security of on-prem and cloud environments, leading incident investigations and response, and more.`,
  
  "Chris": `Chris Williams MA CISSP CCSK SSCP CompTIA Sec + has over 15 years of IT and cybersecurity experience. He works with organizations to bring a better understanding of their security and align security with business objectives. He has provided assessments and advisory services across several sectors including educational, non-profit, manufacturing, and more.`,
  
  "JP": `Jacob Padden Cybersecurity Engineer I has more than five years of experience within the Cyber Security industry. He specializes in Cloud Security, Endpoint Detection, Incident Response, and more. Jacob holds an undergraduate degree in Computer Science with a concentration in Cyber Security from Reinhardt University in Georgia.`,
  
  "Michael": `Michael Nouguier Chief Information Security Officer & Director of Cybersecurity Services has more than 15 years of experience in Information Technology and Cybersecurity. His expertise focuses on enterprise information security and risk management across various industries including financial services and healthcare.`,
  
  "Nicholas": `Nicholas Runco Cybersecurity Engineer has over three years of experience in cybersecurity with certifications in Penetration Testing. He has gained valuable experience at Richey May working on penetration testing, phishing analysis, vulnerability scanning, and incident response.`,
  
  "Parker": `Parker Brissette is an experienced cybersecurity leader with nearly 20 years of experience in cybersecurity and technology practices. He has extensive experience in Fintech, Telecom, Healthcare, and other industries. He holds a Bachelor's and Master’s degree in Cybersecurity and numerous certifications including CISSP and CEH.`,
  
  "Sean": `Sean Kelly-McGeehan Sec+ Net+ Cybersecurity Engineer has more than four years of experience in cybersecurity. Prior to Richey May, Sean worked as a System Administrator and Cybersecurity Administrator in the public education sector. He now supports clients with assessments, phishing campaigns, incident response, and more.`
};

// Function to apply custom styles to the document
async function applyCustomStyles() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    if (paragraphs.items.length > 0) {
      // Apply primary header style
      paragraphs.items[0].font.set({
        name: "Source Sans Pro Black",
        size: 40,
        color: "#002B49",  // Midnight Blue
        bold: true
      });
      paragraphs.items[0].text = paragraphs.items[0].text.toUpperCase();

      // Apply sub-header style
      if (paragraphs.items.length > 1) {
        paragraphs.items[1].font.set({
          name: "Source Sans Pro Bold",
          size: 14,
          color: "#002B49",  // Midnight Blue
          bold: true
        });
        paragraphs.items[1].text = paragraphs.items[1].text.toUpperCase();
      }

      // Apply body text style
      for (let i = 2; i < paragraphs.items.length; i++) {
        paragraphs.items[i].font.set({
          name: "Montserrat",
          size: 10,
          color: "#000000"  // Black
        });
      }
    }

    await context.sync();
  });
}

async function insertTitlePage() {
  await Word.run(async (context) => {
    const body = context.document.body;

    // Function to insert an image using base64 string
    async function insertImage(base64String, width, height) {
      const imageOoxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
          <pkg:xmlData>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                        <wp:extent cx="${width * 9525}" cy="${height * 9525}"/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:blipFill>
                                <a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                  <a:fillRect/>
                                </a:stretch>
                              </pic:blipFill>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
          </pkg:xmlData>
        </pkg:part>
        <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
          <pkg:binaryData>${base64String}</pkg:binaryData>
        </pkg:part>
      </pkg:package>`;

      body.insertOoxml(imageOoxml, Word.InsertLocation.start);
    }

    // Insert mountain image
    const mountainImageBase64 = await fetchImageAsBase64("https://thestateofcybersecurity.github.io/word-addin/assets/mountains.png");
    await insertImage(mountainImageBase64, 600, 400);

    // Insert Richey May logo
    const logoImageBase64 = await fetchImageAsBase64("https://thestateofcybersecurity.github.io/word-addin/assets/logo.png");
    await insertImage(logoImageBase64, 150, 50);

    // Insert text for the title page
    const title = body.insertParagraph("Maturity Assessment", Word.InsertLocation.after);
    title.font.set({
      name: "Source Sans Pro Black",
      size: 40,
      color: "#002B49",  // Midnight Blue
      bold: true
    });
    title.alignment = Word.Alignment.center;

    const preparedFor = body.insertParagraph("Prepared for:", Word.InsertLocation.after);
    preparedFor.font.set({
      name: "Montserrat",
      size: 12,
      color: "#002B49",  // Midnight Blue
    });
    preparedFor.alignment = Word.Alignment.center;

    const date = body.insertParagraph("DATE", Word.InsertLocation.after);
    date.font.set({
      name: "Montserrat",
      size: 12,
      color: "#002B49",  // Midnight Blue
    });
    date.alignment = Word.Alignment.left;

    const deliveredBy = body.insertParagraph("Delivered By: NAME, TITLE", Word.InsertLocation.after);
    deliveredBy.font.set({
      name: "Montserrat",
      size: 12,
      color: "#002B49",  // Midnight Blue
    });
    deliveredBy.alignment = Word.Alignment.left;

    // Insert footer image
    const footerImageBase64 = await fetchImageAsBase64("https://thestateofcybersecurity.github.io/word-addin/assets/greenfooter.png");
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    if (sections.items.length > 0) {
      const firstSection = sections.items[0];
      const footer = firstSection.getFooter(Word.HeaderFooterType.primary);
      footer.insertOoxml(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
          <pkg:xmlData>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                        <wp:extent cx="${600 * 9525}" cy="${100 * 9525}"/>
                        <wp:docPr id="1" name="Picture 1"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:blipFill>
                                <a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                <a:stretch>
                                  <a:fillRect/>
                                </a:stretch>
                              </pic:blipFill>
                            </pic:pic>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
          </pkg:xmlData>
        </pkg:part>
        <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
          <pkg:binaryData>${footerImageBase64}</pkg:binaryData>
        </pkg:part>
      </pkg:package>`, Word.InsertLocation.start);
    }

    await context.sync();
  });
}

// Helper function to fetch image as base64
async function fetchImageAsBase64(url) {
  const response = await fetch(url);
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function insertHeaderFooter() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // Apply the footer only on pages after the first
    sections.items.forEach((section, index) => {
      if (index > 0) {  // Skip the first page (index 0)
        const footer = section.getFooter(Word.HeaderFooterType.primary);

        // Insert horizontal bar and text for footer
        footer.insertParagraph("______________________________________________________________", Word.InsertLocation.start).font.color = "#6AA339";
        const richMayText = footer.insertParagraph("Richey May Cyber", Word.InsertLocation.start);
        const confidentialText = footer.insertParagraph("Confidential", Word.InsertLocation.start);
        const pageNumberText = footer.insertParagraph("Page | ", Word.InsertLocation.start);

        // Insert page number field
        pageNumberText.insertField(Word.FieldType.page, true);

        // Style the footer text
        [richMayText, confidentialText, pageNumberText].forEach((p) => p.font.color = "#6AA339");

        // Align footer text
        richMayText.alignment = Word.Alignment.left;
        confidentialText.alignment = Word.Alignment.center;
        pageNumberText.alignment = Word.Alignment.right;
      }
    });

    await context.sync();
  });
}

// Function to generate table of contents
async function insertTableOfContents() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText("Table of Contents will be inserted here.", Word.InsertLocation.replace);
    range.font.set({
      name: "Montserrat",
      size: 14,
      bold: true
    });
    await context.sync();
  });
}

// Function to insert Excel table (placeholder for actual implementation)
async function insertExcelTable() {
  // Placeholder function for Excel table import
}

// Function to add the selected employee bio into the document
async function addEngagementTeamBios() {
  const selectedMember = document.getElementById("teamDropdown").value;
  const bioText = bios[selectedMember];

  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph(`Bio for ${selectedMember}: ${bioText}`, Word.InsertLocation.end);
    await context.sync();
  });
}
