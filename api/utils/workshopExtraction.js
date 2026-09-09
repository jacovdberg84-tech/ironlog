import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
const exec = promisify(execFile);
export async function extractPages(file, run = exec) {
  const opts = {windowsHide:true,timeout:120000,maxBuffer:32*1024*1024};
  const {stdout} = await run(process.env.PDFTOTEXT_BIN || 'pdftotext', ['-layout','-enc','UTF-8',file,'-'], opts);
  const texts = stdout.split('\f'); if (!texts.at(-1)?.trim()) texts.pop();
  if (!texts.length || texts.length > 2000) throw new Error('Document must contain 1–2000 pages. Split larger manuals.');
  const pages=[]; let ocrError='';
  const tmp=await fs.mkdtemp(path.join(os.tmpdir(),'ironlog-ocr-'));
  try {
    for(let i=0;i<texts.length;i++) {
      let text=texts[i].trim(), method='text', issue='';
      if(text.replace(/\s/g,'').length < 40) {
        if (!ocrError) {
          try {
            const prefix=path.join(tmp,'page');
            await run(process.env.PDFTOPPM_BIN || 'pdftoppm',['-f',String(i+1),'-l',String(i+1),'-scale-to','2400','-singlefile','-png',file,prefix],opts);
            const result=await run(process.env.TESSERACT_BIN || 'tesseract',[prefix+'.png','stdout','-l',process.env.WORKSHOP_OCR_LANG || 'eng'],opts);
            text=result.stdout.trim();method='ocr';
          } catch(err) {
            issue='OCR failed or is unavailable; inspect this page manually.';
            if(err.code==='ENOENT') ocrError=issue;
          }
        } else issue=ocrError;
        if(text.replace(/\s/g,'').length<40) issue ||= 'Little or no readable text; may be a diagram or scan.';
      }
      pages.push({page:i+1,text:text.slice(0,60000),method,issue: text.length>60000 ? 'Page text exceeds indexing limit.' : issue});
    }
    return pages;
  } finally {await fs.rm(tmp,{recursive:true,force:true});}
}
