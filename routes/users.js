let express = require('express');
let router = express.Router();
let extractor = require('textract');
let pptService = require('../services/ppt');
let dropboxService = require('../services/dropbox.service');

/* GET users listing. */
router.post('/', (req, res, next) => {
  extractor.fromBufferWithMime("application/vnd.openxmlformats-officedocument.wordprocessingml.document", Buffer.from(req.body.buffer, 'base64'), {preserveLineBreaks: true}, (error, content) => {
    if (error) {
      console.error(error);
      res.status(400);
      res.json(error);
    } else {
      let annContent = [];
      let ann = content.split('\n');
      let enAndFrAnnouncement = splitEnglishAndFrench(ann);
      Object.entries(enAndFrAnnouncement).forEach(a => {
        annContent.push(announcementContent(a[1]))
      });

      pptService.genPPT(annContent, (buffer) => {
        dropboxService.uploadToDropbox(buffer)
          .then(_ => {
            res.json("L'annonce a été uploadée sur le dropbox avec succès")
          })
          .catch(err => {
            console.error(err);
            res.json("Une erreur est survenue lors de l'upload sur le dropbox");
          });
      });
    }
  })
});

function splitEnglishAndFrench(annonces) {
  let fr = [];
  let en = [];
  let fillFrench = true;
  for (let i = 0; i < annonces.length; i++) {
    let line = annonces[i];
    if (line.toLowerCase().trim().indexOf("the prophetic theme") !== -1) {
      fillFrench = false;
    }

    if (fillFrench) {
      fr.push(line);
    } else {
      en.push(line);
    }
  }

  return {
    fr: fr,
    en: en,
  }
}

function announcementContent(announcement) {
  let annonceContent = {};
  let annonceContentMode = false;
  for (let i = 0; i < announcement.length; i++) {
    let line = announcement[i].toLowerCase().trim();

    if (exclusions.includes(line)) continue;
    if (line.indexOf("thème prophétique") !== -1 || line.indexOf("le thème prophétique") !== -1 || line.indexOf("the prophetic theme") !== -1) {
      annonceContent.themeTitle = line.split(':')[0];
      annonceContent.theme = line.split(':')[1] + ":" + line.split(':')[2];
      if(annonceContent.theme === "undefined") {
        annonceContent.theme = announcement[i + 1];
      }
    }

    if (line.trim() !== "" && annonceContentMode === true) {
      if (annonceContent.content === undefined) annonceContent.content = [];
      if(i == 0) {
        annonceContent.content.push(line);
      } else {
        if(announcement[i - 1].trim() === "") {
          annonceContent.content.push(line);
        } else {
          annonceContent.content[annonceContent.content.length - 1] += line; 
        }
      }
    }

    if (line.indexOf("aujourd'hui est notre") !== -1 || line.indexOf("today is our") !== -1) {
      annonceContent.today = line;
      annonceContentMode = true;
    }
  }

  return annonceContent;
}

const exclusions = [
  "chapelle des vanqueurs internationale paris"
];

module.exports = router;
