const fs=require('node:fs');
const esbuild=require('../installation-video/node_modules/esbuild');
esbuild.buildSync({entryPoints:['tools/password-crypto-entry.mjs'],bundle:true,format:'iife',globalName:'PasswordCrypto',target:'es2020',platform:'browser',outfile:'PasswordCrypto.gs',legalComments:'inline',banner:{js:'// Generated from @noble/hashes 2.0.1 (MIT). Rebuild with node tools/build-password-crypto.cjs.'}});
fs.copyFileSync('node_modules/@noble/hashes/LICENSE','docs/password-crypto-LICENSE.txt');
