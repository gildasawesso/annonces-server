let pptgenerator = require("pptxgenjs");
let path = require('path');
let moment = require('moment');

async function buildpowerpoint(announcement, cb) {
  let pptx = new pptgenerator();
  pptx.setLayout('LAYOUT_WIDE');

  configureMasterSlide(pptx);

  addOpening(pptx);
  addTheme(announcement, pptx);
  addToday(announcement, pptx);
  addAnnoncementsContent(announcement, pptx);

  pptx.save("jszip", (buffer) => cb(buffer), 'nodebuffer');
}

function addTheme(ann, pptx) {
  let slide = pptx.addNewSlide("MASTER");
  slide.addText(ann[0].themeTitle, {x: 1.7, y: 0.5, w: '25%', h: 0.8, color: 'ffffff', fill: '3F51B5', autoFit: true, align: 'center'});
  slide.addText(ann[1].themeTitle, {x: 8.21, y: 0.5, w: '25%', h: 0.8, color: 'ffffff', fill: 'FF9800', autoFit: true, align: 'center'});

  slide.addText(ann[0].theme, {x: 0.26, y: 1.77, w: 6.21, h: 5.43, color: generalOptions.color.fr, ...themOptions});
  slide.addText(ann[0].theme, {x: 6.77, y: 1.77, w: 6.21, h: 5.43, color: generalOptions.color.en, ...themOptions});
}

function addToday(ann, pptx) {
  let slide = pptx.addNewSlide("MASTER_CONTENT");
  slide.addText("Aujourd'hui", specificOptions.frTitleOptions);
  slide.addText("Today", specificOptions.enTitleOptions);
  slide.addText(ann[0].today, specificOptions.frContentOptions);
  slide.addText(ann[1].today, specificOptions.enContentOptions);
}

function addOpening(pptx) {
  let slide = pptx.addNewSlide("MASTER");
  slide.addText("Annonces", {x: 2.07, y: 2.08, w: 9.08, h: 0.67, color: 'ffffff', fill: '3F51B5', autoFit: true, align: 'center'});
  slide.addText("Announcements", {x: 2.07, y: 4.36, w: 9.08, h: 0.67, color: 'ffffff', fill: 'FF9800', autoFit: true, align: 'center'});
}

function addAnnoncementsContent(ann, pptx) {
  for(let i = 0; i < ann[0].content.length; i++) {
    let frAnnounce = ann[0].content[i];
    let enAnnounce = ann[1].content[i];

    let slide = pptx.addNewSlide("MASTER_CONTENT");
    slide.addText("", specificOptions.frTitleOptions);
    slide.addText("", specificOptions.enTitleOptions);
    slide.addText(frAnnounce, specificOptions.frContentOptions);
    slide.addText(enAnnounce, specificOptions.enContentOptions);
  }
}

function configureMasterSlide(pptx) {
  let slideMasterOptions = {
    title: 'MASTER',
    bkgd: {path: `${path.resolve()}/services/bg1.jpeg`}
  };

  let contentSlideMasterOptions = {
    title: 'MASTER_CONTENT',
    bkgd: {path: `${path.resolve()}/services/bg1.jpeg`},
    objects: [
      {'placeholder': {options: {name: 'frTitle', type: 'body', x: 0.5, y: 0.12, w: 4.82, h: 0.75}}},
      {'placeholder': {options: {name: 'enTitle', type: 'body', x: 8, y: 0.12, w: 4.5, h: 0.75}}},
      {'placeholder': {options: {name: 'frContent', type: 'body', x: 0.15, y: 1.06, w: 5.93, h: 6.44}}},
      {'placeholder': {options: {name: 'enContent', type: 'body', x: 7.14, y: 1.06, w: 5.93, h: 6.44}}}
    ]
  };
  pptx.defineSlideMaster(slideMasterOptions);
  pptx.defineSlideMaster(contentSlideMasterOptions);
}

const generalOptions = {
  color: {
    fr: '3F51B5',
    en: 'FF9800'
  },
  common: {
    align: 'center',
    fill: {
      type:'solid',
      color:'000000',
      alpha: 40
    },
    isTextBox: true,
    shrinkText: true,
    fontSize: 36
  }
};

const themOptions = {
  fontSize: 90,
  ...generalOptions.common
};

const specificOptions = {
  frTitleOptions: {
    placeholder: 'frTitle',
    color: 'ffffff',
    fill: generalOptions.color.fr,
    ...generalOptions.common
  },
  enTitleOptions: {
    placeholder: 'enTitle',
    color: 'ffffff',
    fill: generalOptions.color.en,
    ...generalOptions.common
  },
  frContentOptions: {
    placeholder: 'frContent',
    color: '3F51B5',
    ...generalOptions.common
  },
  enContentOptions: {
    placeholder: 'enContent',
    color: 'FF9800',
    ...generalOptions.common
  }
};

exports.genPPT = buildpowerpoint;
