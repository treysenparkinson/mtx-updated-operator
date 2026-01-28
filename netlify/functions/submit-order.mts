import type { Context, Config } from "@netlify/functions";
import * as XLSX from 'xlsx';
import { S3Client, PutObjectCommand } from '@aws-sdk/client-s3';
import PDFDocument from 'pdfkit';

export default async (req: Request, context: Context) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204 });
  }

  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  try {
    const data = await req.json();
    const { refId, contactName, contactEmail, labels } = data;

    if (!refId || !labels || labels.length === 0) {
      return new Response(JSON.stringify({ error: "Missing required fields" }), { status: 400 });
    }

    // Initialize S3 client
    const s3Client = new S3Client({
      region: Netlify.env.get("MY_AWS_REGION") || "us-east-1",
      credentials: {
        accessKeyId: Netlify.env.get("MY_AWS_ACCESS_KEY_ID") || "",
        secretAccessKey: Netlify.env.get("MY_AWS_SECRET_ACCESS_KEY") || ""
      }
    });
    const bucketName = Netlify.env.get("S3_BUCKET") || "";

    const timestamp = new Date().toISOString();
    const formattedDate = new Date().toLocaleString("en-US", {
      year: "numeric", month: "short", day: "2-digit",
      hour: "2-digit", minute: "2-digit"
    });

    // ============================================
    // BUILD EXCEL FILE
    // ============================================
    const excelData: any[][] = [];
    
    // Row 1: Reference ID
    excelData.push([`Reference ID: ${refId}`]);
    
    // Row 2: Headers
    const headers = ["Size", "Color", "VAR1", "VAR2", "VAR3", "VAR4", "VAR5", "VAR6", "VAR1 Size", "VAR2 Size", "VAR3 Size", "VAR4 Size", "VAR5 Size", "VAR6 Size"];
    excelData.push(headers);
    
    let totalLabels = 0;
    labels.forEach((label: any) => {
      const row = [
        label.size?.name || "",
        label.color?.name || "",
        label.var1 || "",
        label.var2 || "",
        label.var3 || "",
        label.var4 || "",
        label.var5 || "",
        label.var6 || "",
        label.var1 ? (label.var1Size || 18) : "",
        label.var2 ? (label.var2Size || 18) : "",
        label.var3 ? (label.var3Size || 18) : "",
        label.var4 ? (label.var4Size || 10) : "",
        label.var5 ? (label.var5Size || 10) : "",
        label.var6 ? (label.var6Size || 10) : ""
      ];
      
      const qty = label.quantity || 1;
      totalLabels += qty;
      for (let i = 0; i < qty; i++) {
        excelData.push(row);
      }
    });

    // Create XLSX workbook
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    
    // Set column widths
    worksheet['!cols'] = [
      { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
      { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 10 }, { wch: 10 },
      { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    
    // Generate XLSX as buffer
    const xlsxBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // ============================================
    // BUILD PDF FILE
    // ============================================
    const pdfBuffer = await generatePDF(refId, formattedDate, labels, totalLabels);

    // ============================================
    // UPLOAD TO S3
    // ============================================
    
    // Upload XLSX to S3
    const xlsxKey = `labels/${refId}/labels-${refId}.xlsx`;
    await s3Client.send(new PutObjectCommand({
      Bucket: bucketName,
      Key: xlsxKey,
      Body: xlsxBuffer,
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      ACL: 'public-read'
    }));
    const xlsxUrl = `https://${bucketName}.s3.${Netlify.env.get("MY_AWS_REGION")}.amazonaws.com/${xlsxKey}`;

    // Upload PDF to S3
    const pdfKey = `labels/${refId}/labels-${refId}.pdf`;
    await s3Client.send(new PutObjectCommand({
      Bucket: bucketName,
      Key: pdfKey,
      Body: pdfBuffer,
      ContentType: 'application/pdf',
      ACL: 'public-read'
    }));
    const pdfUrl = `https://${bucketName}.s3.${Netlify.env.get("MY_AWS_REGION")}.amazonaws.com/${pdfKey}`;

    // ============================================
    // SEND TO ZAPIER WEBHOOK
    // ============================================
    const labelSummaries = labels.map((label: any) => ({
      size: label.size?.name,
      dimensions: label.size?.dimensions,
      color: label.color?.name,
      var1: label.var1 || "",
      var2: label.var2 || "",
      var3: label.var3 || "",
      var4: label.var4 || "",
      var5: label.var5 || "",
      var6: label.var6 || "",
      font: label.font?.name,
      corners: label.corners,
      notch: label.notch,
      quantity: label.quantity
    }));

    const webhookUrl = Netlify.env.get("ZAPIER_WEBHOOK_URL") || "";
    
    const webhookPayload = {
      refId,
      contactName: contactName || "",
      contactEmail: contactEmail || "",
      timestamp,
      formattedDate,
      totalLabels,
      labelCount: labels.length,
      xlsxUrl,
      xlsxFileName: `labels-${refId}.xlsx`,
      pdfUrl,
      pdfFileName: `labels-${refId}.pdf`,
      labels: labelSummaries
    };

    const webhookResponse = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(webhookPayload)
    });

    if (!webhookResponse.ok) {
      console.error("Webhook failed:", await webhookResponse.text());
    }

    return new Response(JSON.stringify({ 
      success: true, 
      message: "Order submitted successfully",
      refId,
      totalLabels,
      pdfUrl,
      xlsxUrl
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });

  } catch (error) {
    console.error("Error processing order:", error);
    return new Response(JSON.stringify({ error: "Failed to process order" }), { status: 500 });
  }
};

// ============================================
// PDF GENERATION FUNCTION
// ============================================
async function generatePDF(refId: string, formattedDate: string, labels: any[], totalLabels: number): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    
    const doc = new PDFDocument({
      size: 'LETTER',
      margins: { top: 50, bottom: 50, left: 50, right: 50 }
    });

    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const pageWidth = 612; // Letter width in points
    const contentWidth = pageWidth - 100; // Minus margins
    
    // Header
    doc.fontSize(12).font('Helvetica-Bold')
       .text(`Reference ID: ${refId}`, 50, 50, { continued: false });
    
    doc.fontSize(10).font('Helvetica')
       .text(formattedDate, 50, 50, { align: 'right' });

    // Title
    doc.moveDown(0.5);
    doc.fontSize(24).font('Helvetica-Bold')
       .text('Saved Labels Summary', 50, doc.y, { continued: false });
    
    // Page number placeholder (we'll handle multi-page later if needed)
    doc.fontSize(10).font('Helvetica')
       .text('Page 1 of 1', 50, 85, { align: 'right' });

    // Divider line
    doc.moveTo(50, 110).lineTo(pageWidth - 50, 110).stroke('#e2e8f0');

    let yPosition = 130;
    const labelHeight = 100;
    const labelPreviewWidth = 70;
    const labelPreviewHeight = 80;

    labels.forEach((label, index) => {
      // Check if we need a new page
      if (yPosition + labelHeight > 720) {
        doc.addPage();
        yPosition = 50;
      }

      const size = label.size || { width: 160, height: 182, name: "30MM Standard", dimensions: '2" × 2.27"' };
      const color = label.color || { bg: "#16a34a", text: "#fff", name: "Green/White" };
      const corners = label.corners || "squared";

      // Draw label preview
      const previewX = 50;
      const previewY = yPosition + 10;
      
      // Scale to fit preview area
      const scaleX = labelPreviewWidth / size.width;
      const scaleY = labelPreviewHeight / size.height;
      const scale = Math.min(scaleX, scaleY);
      const scaledWidth = size.width * scale;
      const scaledHeight = size.height * scale;
      
      // Center the preview
      const offsetX = previewX + (labelPreviewWidth - scaledWidth) / 2;
      const offsetY = previewY + (labelPreviewHeight - scaledHeight) / 2;

      // Draw label background
      const borderRadius = corners === 'rounded' ? 5 : 0;
      doc.roundedRect(offsetX, offsetY, scaledWidth, scaledHeight, borderRadius)
         .fill(color.bg);

      // Draw cutout circle
      const cutoutR = (size.id === '22mm' ? 24 : 36) * scale;
      const cutoutY = offsetY + scaledHeight * 0.68;
      const cutoutX = offsetX + scaledWidth / 2;
      doc.circle(cutoutX, cutoutY, cutoutR).fill('#ffffff');

      // Draw text on label (VAR1 only for preview)
      if (label.var1) {
        const textColor = color.text || '#ffffff';
        doc.fontSize(8 * scale).font('Helvetica-Bold').fillColor(textColor)
           .text(label.var1, offsetX, offsetY + 8, { 
             width: scaledWidth, 
             align: 'center' 
           });
      }

      // Reset fill color for rest of document
      doc.fillColor('#000000');

      // Label info (to the right of preview)
      const infoX = 140;
      
      doc.fontSize(14).font('Helvetica-Bold')
         .text(size.name, infoX, yPosition + 20);
      
      doc.fontSize(11).font('Helvetica').fillColor('#64748b')
         .text(size.dimensions, infoX, yPosition + 40);

      doc.fillColor('#000000');

      // Quantity (far right)
      doc.fontSize(16).font('Helvetica-Bold')
         .text(`×${label.quantity || 1}`, pageWidth - 100, yPosition + 35, { align: 'right' });

      // Divider line
      yPosition += labelHeight;
      doc.moveTo(50, yPosition).lineTo(pageWidth - 50, yPosition).stroke('#e2e8f0');
      
      yPosition += 10;
    });

    doc.end();
  });
}

export const config: Config = {
  path: "/api/submit-order"
};
