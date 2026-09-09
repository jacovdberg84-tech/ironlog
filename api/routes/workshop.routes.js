import { createWorkshopIndexer, answerWorkshop } from '../utils/workshopKnowledge.js';
import multipart from '@fastify/multipart';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { pipeline } from 'node:stream/promises';
import { Transform } from 'node:stream';
import { getDataRoot } from '../utils/storagePaths.js';
import { isAuthRequired } from '../auth/config.js';

export default async function workshopRoutes(app, options = {}) {
  const db = options.db || (await import('../db/client.js')).db;
  const root = path.join(options.dataRoot || getDataRoot(), 'workshop-files');
  await fsp.mkdir(root, { recursive: true });
  await app.register(multipart, { limits: { fileSize: 100 * 1024 * 1024, files: 1, fields: 8, fieldSize: 2000 } });
  db.exec(`CREATE TABLE IF NOT EXISTS workshop_documents (
    id TEXT PRIMARY KEY, title TEXT NOT NULL, doc_type TEXT NOT NULL,
    manufacturer TEXT NOT NULL DEFAULT '', model TEXT NOT NULL DEFAULT '',
    revision TEXT NOT NULL DEFAULT '', applicability TEXT NOT NULL DEFAULT '',
    filename TEXT NOT NULL, size_bytes INTEGER NOT NULL, uploaded_by TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
  )`);
  const indexer = createWorkshopIndexer(db, root, path);
  const types = new Set(['Parts Manual','Workshop Manual','Service Manual','OEM Bulletin','Technical Document','Workshop Fix']);
  app.get('/documents', async req => {
    const q = String(req.query?.q || '').trim().slice(0,200);
    const rows = db.prepare(`SELECT * FROM workshop_documents WHERE
      instr(lower(title || ' ' || manufacturer || ' ' || model || ' ' || doc_type || ' ' || revision || ' ' || applicability), lower(?)) > 0
      ORDER BY created_at DESC, id DESC LIMIT 500`).all(q);
    return {ok:true, documents:rows.map(r=>({...r,index:db.prepare('SELECT * FROM workshop_index_jobs WHERE document_id=?').get(r.id)||{status:'not_indexed'}})), limit:500};
  });
  app.post('/documents/:id/index', async (req,reply) => {
    const roles=(String(req.headers['x-user-roles']||'')+','+String(req.headers['x-user-role']||'')).split(',').map(s=>s.trim());
    if (!roles.some(r=>['admin','supervisor','workshop_admin','plant_manager'].includes(r)) && (isAuthRequired() || roles.some(Boolean))) return reply.code(403).send({error:'Workshop administrator or supervisor required.'});
    const id=String(req.params.id);
    if(!/^[a-f0-9-]{36}$/.test(id) || !db.prepare('SELECT id FROM workshop_documents WHERE id=?').get(id))return reply.code(404).send({error:'Document not found'});
    indexer.enqueue(id);return reply.code(202).send({ok:true});
  });
  app.post('/ask', async (req,reply) => {
    const question=String(req.body?.question||'').trim();
    if(!question || question.length>2000)return reply.code(400).send({error:'Enter a question up to 2000 characters.'});
    return answerWorkshop(db,question,String(req.body?.document_id||''));
  });
  app.post('/documents', async (req, reply) => {
    const roles = (String(req.headers['x-user-roles'] || '') + ',' + String(req.headers['x-user-role'] || '')).split(',').map(s=>s.trim());
    if (!roles.some(r=>['admin','supervisor','workshop_admin','plant_manager'].includes(r)) && (isAuthRequired() || roles.some(Boolean))) {
      return reply.code(403).send({error:'A workshop administrator or supervisor must upload documents.'});
    }
    const id = randomUUID(), target = path.join(root, id + '.pdf');
    let saved = false, fileSeen = false, size = 0, header = Buffer.alloc(0), filename = '';
    const fields = {};
    try {
      for await (const part of req.parts()) {
        if (part.type !== 'file') {
          if (part.valueTruncated) throw Object.assign(new Error('Document details are too long.'), {statusCode:400});
          fields[part.fieldname] = String(part.value || '').trim();
          continue;
        }
        fileSeen = true;
        filename = path.basename(String(part.filename || '').replaceAll('\\','/')).replace(/[\r\n\x00-\x1f]/g,'').slice(0,180);
        if (!filename.toLowerCase().endsWith('.pdf')) throw Object.assign(new Error('Please upload a PDF document.'), {statusCode:400});
        const inspect = new Transform({transform(chunk, _enc, done) {
          size += chunk.length;
          if (header.length < 5) header = Buffer.concat([header, chunk]).subarray(0,5);
          done(null, chunk);
        }});
        await pipeline(part.file, inspect, fs.createWriteStream(target, {flags:'wx'}));
        if (part.file.truncated) throw Object.assign(new Error('PDF exceeds the 100 MB limit.'), {statusCode:413});
      }
      if (!fileSeen || header.toString() !== '%PDF-') throw Object.assign(new Error('A valid PDF file is required.'), {statusCode:400});
      const title = fields.title || filename.replace(/\.pdf$/i,'');
      if (title.length > 200 || !types.has(fields.doc_type)) throw Object.assign(new Error('Provide a title up to 200 characters and a valid document type.'), {statusCode:400});
      for (const key of ['manufacturer','model','revision','applicability']) if ((fields[key] || '').length > 500) throw Object.assign(new Error('Document details exceed 500 characters.'), {statusCode:400});
      db.prepare(`INSERT INTO workshop_documents (id,title,doc_type,manufacturer,model,revision,applicability,filename,size_bytes,uploaded_by)
        VALUES (?,?,?,?,?,?,?,?,?,?)`).run(id,title,fields.doc_type,fields.manufacturer||'',fields.model||'',fields.revision||'',fields.applicability||'',filename,size,String(req.headers['x-user-name'] || 'local-user'));
      saved = true;
      return reply.code(201).send({ok:true,id});
    } catch (err) {
      await fsp.rm(target, {force:true});
      req.log.error(err);
      return reply.code(err.statusCode || 500).send({error:err.statusCode && err.statusCode < 500 ? err.message : 'Upload failed. Please retry.'});
    } finally {
      if (!saved) await fsp.rm(target, {force:true});
    }
  });
  app.get('/documents/:id/file', async (req,reply) => {
    const row = db.prepare('SELECT * FROM workshop_documents WHERE id=?').get(String(req.params.id));
    if (!row || !/^[a-f0-9-]{36}$/.test(row.id)) return reply.code(404).send({error:'Document not found'});
    const file = path.join(root,row.id + '.pdf');
    try { await fsp.access(file); } catch { return reply.code(404).send({error:'Document file unavailable'}); }
    return reply.header('Cache-Control','no-store').header('X-Content-Type-Options','nosniff')
      .header('Content-Disposition',"attachment; filename*=UTF-8''" + encodeURIComponent(row.filename))
      .type('application/pdf').send(fs.createReadStream(file));
  });
}

