import * as React from "react";
import { useEffect, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";

// Set up worker
pdfjsLib.GlobalWorkerOptions.workerSrc = require("pdfjs-dist/build/pdf.worker.entry");

interface PdfThumbnailProps {
  fileUrl: string;
  width?: number;
}
// Set up worker - Using a version-matched CDN for reliable loading in SPFx
const PDFJS_VERSION = '2.16.105';
// We avoid setting workerSrc to a CDN to prevent CSP violations. 
// Instead, we use disableWorker: true in the getDocument options.

const PdfThumbnail: React.FC<PdfThumbnailProps> = ({ fileUrl, width }) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [status, setStatus] = useState<'loading' | 'success' | 'error'>('loading');

  useEffect(() => {
    const renderThumbnail = async () => {
      const loadingTask = pdfjsLib.getDocument(fileUrl);
      const pdf = await loadingTask.promise;
      const page = await pdf.getPage(1); // First page only

      const viewport = page.getViewport({ scale: 1 });
      const scale = width / viewport.width; // resize to thumbnail width
      const scaledViewport = page.getViewport({ scale });

      const canvas = canvasRef.current!;
      const context = canvas.getContext("2d")!;
      canvas.width = scaledViewport.width;
      canvas.height = scaledViewport.height;

      await page.render({ canvasContext: context, viewport: scaledViewport }).promise;
    };

    renderThumbnail();
  }, [fileUrl, width]);

  return (
    <div style={{ padding: "4px", display: "inline-block" }}>
      <canvas ref={canvasRef} style={{ objectFit: 'cover' }} />
    </div>
  );
};

export default PdfThumbnail;
