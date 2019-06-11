let Dropbox = require('dropbox').Dropbox;
let fetch = require('isomorphic-fetch');
let fs = require('fs');
let path = require('path');
let moment = require('moment');

const announcementsFolder = "/Winner Chapel Paris/Annonces/Annonces ppt/2019";

async function uploadToDropbox(buffer) {
  let dbx = new Dropbox({ accessToken: 'D_QAGxU6zJMAAAAAAAAjbHjBRYBw7KEo_4vDc2iKJCZ62tnrtv9pHVELZyTSz37D', fetch: fetch });
  try {
    let filesList = await dbx.filesListFolder({path: announcementsFolder});
    if(filesList.entries.find(f => f.name === fileName()) !== undefined) {
      await dbx.filesDeleteV2({path: fileName()});
    }
    return await dbx.filesUpload({path: fileName(), contents: buffer});
  }catch(err){
    console.error(err);
    return err;
  }
}

function fileName() {
  return `/Winner Chapel Paris/Annonces/Annonces ppt/2019/${moment().day(7).format('DD MMMM YYYY')}.pptx`;
}

exports.uploadToDropbox = uploadToDropbox;
