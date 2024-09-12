/* global Office, Word, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
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
      // Apply primary header style (Source Sans Pro Black, 40pt, Midnight Blue, ALL CAPS)
      const headerText = paragraphs.items[0].text.toUpperCase();  // Transform to uppercase
      paragraphs.items[0].insertText(headerText, Word.InsertLocation.replace);
      paragraphs.items[0].font.set({
        name: "Source Sans Pro Black",
        size: 40,
        color: "#002B49",  // Midnight Blue
        bold: true
      });

      // Apply sub-header style (Source Sans Pro Bold, 12-14pt, Midnight Blue, ALL CAPS)
      if (paragraphs.items.length > 1) {
        const subHeaderText = paragraphs.items[1].text.toUpperCase();  // Transform to uppercase
        paragraphs.items[1].insertText(subHeaderText, Word.InsertLocation.replace);
        paragraphs.items[1].font.set({
          name: "Source Sans Pro Bold",
          size: 14,
          color: "#002B49",  // Midnight Blue
          bold: true,
          allCaps: true,
          kerning: 100
        });
      }

      // Apply body text style (Montserrat, 10pt, black)
      for (let i = 2; i < paragraphs.items.length; i++) {
        paragraphs.items[i].font.set({
          name: "Montserrat",
          size: 10,
          color: "#000000",  // Black
          justify: true
        });
      }
    }

    await context.sync();
  });
}

// Function to insert header and footer
async function insertHeaderFooter() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load();
    await context.sync();

    const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);
    header.insertParagraph("Richey May Confidential", Word.InsertLocation.start);

    const footer = sections.items[0].getFooter(Word.HeaderFooterType.primary);
    footer.insertParagraph("Page 1", Word.InsertLocation.start);

    await context.sync();
  });
}

// Function to generate table of contents
async function insertTableOfContents() {
  await Word.run(async (context) => {
    context.document.body.insertTableOfContents(Word.InsertLocation.start, {
      headingStyleCount: 3,
      rightAlignPageNumbers: true,
      useHyperlinks: true,
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
