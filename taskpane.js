Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Ensure all elements exist before attaching event listeners
    const elements = [
      { id: "insertTitlePageButton", handler: insertTitlePage },
      { id: "applyStylesButton", handler: applyCustomStyles },
      { id: "insertHeaderFooterButton", handler: insertHeaderFooter },
      { id: "generateTOCButton", handler: insertTableOfContents },
      { id: "importExcelTable", handler: insertExcelTable },
      { id: "addBiosButton", handler: addEngagementTeamBios },
      { id: "insertReportTemplateButton", handler: insertReportTemplate }
    ];
    
    elements.forEach(({ id, handler }) => {
      const element = document.getElementById(id);
      if (element) {
        element.onclick = handler;
      } else {
        console.warn(`Element with id "${id}" not found`);
      }
    });
  }
});
// Employee bios data loaded from attached files
const bios = {
  "Alvin": `Alvin Tugume Cybersecurity Consultant: vCISO Infosec. Engineer Incident Responder Risk Assessor CCSKv4 | CompTIA SEC+ | TPN Certified With 10 years in IT and Cybersecurity Alvin Tugume is a recognized expert in the field. A proud holder of a Bachelor's degree in Cybersecurity. Prior to Richey May Alvin was responsible for the cybersecurity posture of three credit unions and a call center. Responsibilities included maintaining compliance with information security policies, monitoring the security of on-prem and cloud environments, leading incident investigations and response, and more.`,
  
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

async function insertTitlePage() {
  await Word.run(async (context) => {
    const document = context.document;
    const body = document.body;

    // Clear existing content
    body.clear();

     // Insert mountain image with specific settings
    const mountainImage = await insertImage(context, "https://thestateofcybersecurity.github.io/word-addin/assets/mountains.png", null, null, Word.InsertLocation.start);
    
    // Set specific positioning and size for the mountain image
    mountainImage.setPositioning({
      layoutWrap: Word.ImageLayoutWrap.behind,
      left: -9.84,
      leftRelative: Word.ImageRelativeHorizontalPosition.page,
      top: -0.026,
      topRelative: Word.ImageRelativeVerticalPosition.paragraph,
      height: 11.83,
      width: 23.55,
      lockAspectRatio: true,
      allowOverlap: true,
    });

    // Set scale (we need to do this separately as it's not part of setPositioning)
    mountainImage.scaleHeight = 360;
    mountainImage.scaleWidth = 361;

    // Ensure the image doesn't move with text
    mountainImage.lockPosition = true;
    
    // Insert company logo with specific settings
    const logoImage = await insertImage(context, "https://thestateofcybersecurity.github.io/word-addin/assets/logo.png", null, null, Word.InsertLocation.end);
    
    logoImage.setPositioning({
      layoutWrap: Word.ImageLayoutWrap.square,
      left: 1.56,
      leftRelative: Word.ImageRelativeHorizontalPosition.margin,
      top: 1.95,
      topRelative: Word.ImageRelativeVerticalPosition.paragraph,
      height: 1.99,
      width: 5.22,
      lockAspectRatio: false,
      allowOverlap: true,
    });
    logoImage.lockPosition = true;

    // Insert blank paragraph for spacing
    body.insertParagraph("", Word.InsertLocation.end);

    // Insert title text
    const title = body.insertParagraph("Maturity Assessment", Word.InsertLocation.end);
    title.font.set({
      name: "Source Sans Pro Black",
      size: 40,
      color: "#002B49",
      bold: true
    });
    title.alignment = Word.Alignment.centered;

    // Insert other title page elements
    const preparedFor = body.insertParagraph("Prepared for: [Client Name]", Word.InsertLocation.end);
    preparedFor.font.set({
      name: "Montserrat",
      size: 12,
      color: "#002B49",
    });
    preparedFor.alignment = Word.Alignment.centered;

    // Insert blank paragraph for spacing
    body.insertParagraph("", Word.InsertLocation.end);

    const date = body.insertParagraph("Date: " + new Date().toLocaleDateString(), Word.InsertLocation.end);
    date.font.set({
      name: "Montserrat",
      size: 12,
      color: "#002B49",
    });
    date.alignment = Word.Alignment.left;

    // Insert footer image in the first page footer
    const sections = document.sections;
    sections.load("items");
    await context.sync();

    const firstPageFooter = sections.items[0].getFooter(Word.HeaderFooterType.firstPage);
    await insertImage(context, "https://thestateofcybersecurity.github.io/word-addin/assets/greenfooter.png", 600, 100, Word.InsertLocation.replace, firstPageFooter);

    // Insert page break
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    await context.sync();
  });
}

async function insertImage(context, url, width, height, location, target = context.document.body) {
  const base64Image = await fetchImageAsBase64(url);
  const image = target.insertInlinePictureFromBase64(base64Image, location);
  
  if (width && height) {
    image.width = width;
    image.height = height;
  }
  
  await context.sync();
  return image;
}

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

    sections.items.forEach((section, index) => {
      if (index > 0) {  // Skip the first page
        const footer = section.getFooter(Word.HeaderFooterType.primary);

        // Insert horizontal bar
        const horizontalLine = footer.insertParagraph("", Word.InsertLocation.start);
        horizontalLine.font.color = "#6AA339";
        horizontalLine.font.size = 14;
        horizontalLine.insertHorizontalLine();

        // Insert footer text
        const footerText = footer.insertParagraph("Richey May Cyber | Confidential | Page ", Word.InsertLocation.end);
        footerText.font.color = "#6AA339";
        footerText.font.size = 10;
        footerText.alignment = Word.Alignment.right;

        // Insert page number
        footerText.insertField(Word.FieldType.page, Word.InsertLocation.end);
      }
    });

    await context.sync();
  });
}

async function insertReportTemplate() {
  const templateSelect = document.getElementById("templateSelect");
  const selectedTemplate = templateSelect.value;

  await Word.run(async (context) => {
    const body = context.document.body;

    switch (selectedTemplate) {
      case "nistcsf":
        // Insert NIST CSF template content
        body.insertText("[NIST CSF Report Template Content]", Word.InsertLocation.end);
        break;
      case "iso27001":
        // Insert ISO 27001 template content
        body.insertText("[ISO 27001 Report Template Content]", Word.InsertLocation.end);
        break;
      // Add more cases for other templates
    }

    await context.sync();
  });
}

async function addEngagementTeamBios() {
  const selectedMember = document.getElementById("teamDropdown").value;
  const bioText = bios[selectedMember];

  await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert bio header
    const bioHeader = body.insertParagraph(selectedMember, Word.InsertLocation.end);
    bioHeader.font.set({
      name: "Source Sans Pro Black",
      size: 16,
      color: "#002B49",
      bold: true
    });

    // Insert bio text
    const bioParagraph = body.insertParagraph(bioText, Word.InsertLocation.end);
    bioParagraph.font.set({
      name: "Montserrat",
      size: 11,
      color: "#000000"
    });

    await context.sync();
  });
}
