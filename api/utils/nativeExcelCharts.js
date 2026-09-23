import JSZip from "jszip";

const CONTENT_TYPES = "[Content_Types].xml";
const WORKBOOK = "xl/workbook.xml";
const WORKBOOK_RELS = "xl/_rels/workbook.xml.rels";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function fileText(zip, fileName) {
  const file = zip.file(fileName);
  if (!file) throw new Error(`Excel chart export could not find ${fileName}`);
  return file.async("string");
}

function nextRelationshipId(xml) {
  const ids = Array.from(String(xml || "").matchAll(/\bId="rId(\d+)"/g))
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  return `rId${Math.max(0, ...ids) + 1}`;
}

function nextPartNumber(zip, folder, stem) {
  const escapedStem = String(stem).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const matcher = new RegExp(`^${folder}/${escapedStem}(\\d+)\\.xml$`);
  const numbers = Object.keys(zip.files)
    .map((name) => name.match(matcher))
    .filter(Boolean)
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  return Math.max(0, ...numbers) + 1;
}

function getWorksheetPath(workbookXml, workbookRelsXml, sheetName) {
  const sheetPattern = /<sheet\b([^>]*)\/?>(?:<\/sheet>)?/g;
  let sheetMatch;
  let workbookRelationshipId = "";
  while ((sheetMatch = sheetPattern.exec(workbookXml))) {
    const attributes = sheetMatch[1] || "";
    const name = attributes.match(/\bname="([^"]*)"/);
    if (name?.[1] !== sheetName) continue;
    workbookRelationshipId = attributes.match(/\br:id="([^"]*)"/)?.[1] || "";
    break;
  }
  if (!workbookRelationshipId) throw new Error(`Excel chart export could not find worksheet “${sheetName}”`);

  const relPattern = /<Relationship\b([^>]*)\/?>(?:<\/Relationship>)?/g;
  let relMatch;
  while ((relMatch = relPattern.exec(workbookRelsXml))) {
    const attributes = relMatch[1] || "";
    if (attributes.match(/\bId="([^"]*)"/)?.[1] !== workbookRelationshipId) continue;
    const target = attributes.match(/\bTarget="([^"]*)"/)?.[1] || "";
    if (!target) break;
    return `xl/${target.replace(/^\//, "")}`.replace(/\/+/g, "/");
  }
  throw new Error(`Excel chart export could not resolve worksheet “${sheetName}”`);
}

function worksheetRelsPath(worksheetPath) {
  const parts = worksheetPath.split("/");
  const filename = parts.pop();
  return `${parts.join("/")}/_rels/${filename}.rels`;
}

function addContentTypeOverride(contentTypesXml, partName, contentType) {
  if (contentTypesXml.includes(`PartName="/${partName}"`)) return contentTypesXml;
  const override = `<Override PartName="/${partName}" ContentType="${contentType}"/>`;
  return contentTypesXml.replace("</Types>", `${override}</Types>`);
}

function relationshipXml(relationshipId, type, target) {
  return `<Relationship Id="${relationshipId}" Type="${type}" Target="${target}"/>`;
}

function asNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function stringCache(values) {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join("")}</c:strCache>`;
}

function numberCache(values) {
  return `<c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${asNumber(value)}</c:v></c:pt>`).join("")}</c:numCache>`;
}

function barChartXml({ chartId, title, seriesName, categoryFormula, valueFormula, categories, values, color = "2F75B5" }) {
  const categoryAxisId = 500000 + chartId * 2;
  const valueAxisId = categoryAxisId + 1;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="17365D"/></a:solidFill></a:rPr><a:t>${escapeXml(title)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx><c:layout/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>${escapeXml(seriesName)}</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>${escapeXml(categoryFormula)}</c:f>${stringCache(categories)}</c:strRef></c:cat>
          <c:val><c:numRef><c:f>${escapeXml(valueFormula)}</c:f>${numberCache(values)}</c:numRef></c:val>
        </c:ser>
        <c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valueAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${valueAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="0.0" sourceLinked="0"/><c:majorGridlines/><c:tickLblPos val="nextTo"/><c:crossAx val="${categoryAxisId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>
</c:chartSpace>`;
}

function drawingAnchor({ index, relationshipId, position, title }) {
  const from = position?.from || { col: 0, row: 0 };
  const to = position?.to || { col: 6, row: 14 };
  return `<xdr:twoCellAnchor editAs="oneCell">
  <xdr:from><xdr:col>${Math.max(0, Number(from.col) || 0)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${Math.max(0, Number(from.row) || 0)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>${Math.max(1, Number(to.col) || 6)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${Math.max(1, Number(to.row) || 14)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 1}" name="${escapeXml(title)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relationshipId}"/></a:graphicData></a:graphic></xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>`;
}

/**
 * ExcelJS writes the workbook data and styles used by Ironlog. It does not write
 * native charts, so this small post-processor adds editable Open XML bar charts
 * without changing any report values or formulas.
 */
export async function addNativeBarCharts(xlsxBuffer, { sheetName, charts = [] } = {}) {
  const activeCharts = charts.filter((chart) => Array.isArray(chart.categories) && chart.categories.length && chart.categories.length === chart.values?.length);
  if (!activeCharts.length) return Buffer.from(xlsxBuffer);

  const zip = await JSZip.loadAsync(xlsxBuffer);
  const [workbookXml, workbookRelsXml, contentTypesXml] = await Promise.all([
    fileText(zip, WORKBOOK),
    fileText(zip, WORKBOOK_RELS),
    fileText(zip, CONTENT_TYPES),
  ]);
  const worksheetPath = getWorksheetPath(workbookXml, workbookRelsXml, String(sheetName || ""));
  const worksheetXml = await fileText(zip, worksheetPath);
  if (worksheetXml.includes("<drawing")) {
    throw new Error(`Excel chart export cannot add charts to worksheet “${sheetName}” because it already contains drawings`);
  }

  const drawingNumber = nextPartNumber(zip, "xl/drawings", "drawing");
  const drawingPath = `xl/drawings/drawing${drawingNumber}.xml`;
  const drawingRelsPath = `xl/drawings/_rels/drawing${drawingNumber}.xml.rels`;
  const sheetRelsPath = worksheetRelsPath(worksheetPath);
  let worksheetRelsXml = zip.file(sheetRelsPath)
    ? await fileText(zip, sheetRelsPath)
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"></Relationships>`;
  const drawingRelationshipId = nextRelationshipId(worksheetRelsXml);
  worksheetRelsXml = worksheetRelsXml.replace("</Relationships>", `${relationshipXml(drawingRelationshipId, DRAWING_REL_TYPE, `../drawings/drawing${drawingNumber}.xml`)}</Relationships>`);

  let contentTypes = addContentTypeOverride(contentTypesXml, drawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml");
  const drawingRelationships = [];
  const anchors = [];
  for (const [index, chart] of activeCharts.entries()) {
    const chartNumber = nextPartNumber(zip, "xl/charts", "chart");
    const chartPath = `xl/charts/chart${chartNumber}.xml`;
    const chartRelationshipId = `rId${index + 1}`;
    contentTypes = addContentTypeOverride(contentTypes, chartPath, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    zip.file(chartPath, barChartXml({ chartId: chartNumber, ...chart }));
    drawingRelationships.push(relationshipXml(chartRelationshipId, CHART_REL_TYPE, `../charts/chart${chartNumber}.xml`));
    anchors.push(drawingAnchor({ index, relationshipId: chartRelationshipId, position: chart.position, title: chart.title }));
  }

  const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors.join("")}</xdr:wsDr>`;
  const drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}">${drawingRelationships.join("")}</Relationships>`;

  zip.file(CONTENT_TYPES, contentTypes);
  zip.file(sheetRelsPath, worksheetRelsXml);
  zip.file(drawingPath, drawingXml);
  zip.file(drawingRelsPath, drawingRelsXml);
  zip.file(worksheetPath, worksheetXml.replace("</worksheet>", `<drawing r:id="${drawingRelationshipId}"/></worksheet>`));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
