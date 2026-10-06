const fs = require('node:fs');
const path = require('node:path');
const {execFileSync} = require('node:child_process');

async function publish() {
  const base = process.env.CATALOG_API_URL;
  const slug = process.env.CATALOG_PRODUCT_SLUG;
  const token = process.env.CATALOG_SYNC_TOKEN;
  if(base !== 'https://catalog.factorystudio.pro' || slug !== 'leave-management-system' || !token) throw new Error('Catalog settings or OIDC token are unavailable');
  const tracked = execFileSync('git',['ls-files'],{encoding:'utf8'}).trim().split('\n');
  if(tracked.some(file => /(^|\/)\.(env|clasp|clasprc|dev\.vars)(\.|$)|\.(pem|p12|pfx|key)$|(^|\/)(credentials|secrets?)\.(json|ya?ml)$/i.test(file))) throw new Error('Secret-like tracked file detected');
  const source = execFileSync('git',['archive','--format=tar','HEAD','Code.gs','Dashboard.html','appsscript.json','README.md','docs','installation-video','package.json','tests'],{maxBuffer:20*1024*1024});
  if(/AKIA[0-9A-Z]{16}|BEGIN (?:RSA |OPENSSH |EC )?PRIVATE KEY/.test(source.toString('latin1'))) throw new Error('Secret-like source content detected');
  const revision = (process.env.GITHUB_SHA || execFileSync('git',['rev-parse','HEAD'],{encoding:'utf8'}).trim()).slice(0,8);
  const version = require('../package.json').version + '+' + revision;
  const endpoint=base+'/api/integrations/github-release';
  const auth={Authorization:'Bearer '+token};
  async function jsonRequest(method,headers,body) {
    const response=await fetch(endpoint,{method,headers:{...auth,...headers},body,signal:AbortSignal.timeout(120000)});
    const data=await response.json();
    if(!response.ok) throw new Error(`Catalog ${method} failed (${response.status}): ${data.error || 'Unknown error'}`);
    return data;
  }
  async function upload(file,type,label,contentType) {
    const bytes=fs.readFileSync(file),name=path.basename(file);
    const headers={'X-Product-Slug':slug,'X-Release-Version':version,'X-Asset-Type':type,'X-Asset-Label':label,'X-File-Name':name,'X-File-Size':String(bytes.length),'Content-Type':contentType};
    const prepared=await jsonRequest('POST',headers);
    if(!prepared.upload?.uploadUrl || !prepared.upload?.key || !prepared.upload?.id) throw new Error('Catalog returned an incomplete upload');
    const url=new URL(prepared.upload.uploadUrl);
    if(url.protocol!=='https:' || !url.hostname.endsWith('.r2.cloudflarestorage.com')) throw new Error('Unexpected upload destination');
    const put=await fetch(url,{method:'PUT',headers:{'Content-Type':contentType},body:bytes,signal:AbortSignal.timeout(300000)});
    if(!put.ok) throw new Error('R2 upload failed: '+put.status);
    const finalized=await jsonRequest('PATCH',{...headers,'X-Upload-Id':prepared.upload.id,'X-Upload-Key':prepared.upload.key});
    if(!finalized.ok) throw new Error('Catalog could not verify uploaded file');
    if(type!=='cover' && (finalized.asset?.version!==version || finalized.asset?.size!==bytes.length || finalized.asset?.fileName!==name)) throw new Error('Uploaded asset verification mismatch');
    console.log(`Verified ${type}: ${name} (${bytes.length} bytes)`);
  }
  await upload('dist/source-code.zip','package','Source code','application/zip');
  for(const name of fs.readdirSync('release-assets').sort()) {
    const file=path.join('release-assets',name);
    if(name.endsWith('.pdf')) await upload(file,'manual_pdf','Thai installation guide','application/pdf');
    if(name.endsWith('.mp4')) await upload(file,'video','Thai landscape installation tutorial','video/mp4');
  }
  if(fs.existsSync('release-assets/cover.png')) await upload('release-assets/cover.png','cover','Leave management cover','image/png');
  const metadata=JSON.parse(fs.readFileSync('catalog-product.json','utf8'));
  if(metadata.slug!==slug) throw new Error('Product metadata slug mismatch');
  metadata.version=version;
  const updated=await jsonRequest('PUT',{'X-Product-Slug':slug,'Content-Type':'application/json'},JSON.stringify(metadata));
  if(!updated.ok || updated.product?.version!==version) throw new Error('Product metadata verification failed');
  if(process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY,`## FactoryStudio Catalog updated\n\nProduct: ${slug}\n\nVersion: ${version}\n\nCatalog: ${base}/catalog\n`);
  console.log('Catalog updated:',slug,version);
}
publish().catch(error=>{console.error(error.message);process.exitCode=1;});
