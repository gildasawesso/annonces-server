let pptgenerator = require("pptxgenjs");
let path = require('path');
let moment = require('moment');
let officegen = require('officegen');

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
  slide.addText(ann[0].themeTitle, {
    x: 1.7,
    y: 0.5,
    w: '25%',
    h: 0.8,
    color: 'ffffff',
    fill: '3F51B5',
    autoFit: true,
    align: 'center'
  });
  slide.addText(ann[1].themeTitle, {
    x: 8.21,
    y: 0.5,
    w: '25%',
    h: 0.8,
    color: 'ffffff',
    fill: 'FF9800',
    autoFit: true,
    align: 'center'
  });

  slide.addText(ann[0].theme, { x: 0.26, y: 1.77, w: 6.21, h: 5.43, color: generalOptions.color.fr, ...themOptions });
  slide.addText(ann[0].theme, { x: 6.77, y: 1.77, w: 6.21, h: 5.43, color: generalOptions.color.en, ...themOptions });
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
  slide.addText("Annonces", {
    x: 2.07,
    y: 2.08,
    w: 9.08,
    h: 0.67,
    color: 'ffffff',
    fill: '3F51B5',
    autoFit: true,
    align: 'center'
  });
  slide.addText("Announcements", {
    x: 2.07,
    y: 4.36,
    w: 9.08,
    h: 0.67,
    color: 'ffffff',
    fill: 'FF9800',
    autoFit: true,
    align: 'center'
  });
}

function addAnnoncementsContent(ann, pptx) {
  for (let i = 0; i < ann[0].content.length; i++) {
    const regex = /\*.*?\*/g;
    let frAnnounce;
    let enAnnounce;
    try {
      frAnnounce = ann[0].content[i].match(regex).toString().replace(/\*/g, '');
    } catch (e) {
      frAnnounce = "";
    }

    try {
      enAnnounce = ann[1].content[i].match(regex).toString().replace(/\*/g, '');
    } catch (e) {
      enAnnounce = "";
    }



    let slide = pptx.addNewSlide("MASTER_CONTENT");
    slide.addText("", specificOptions.frTitleOptions);
    slide.addText("", specificOptions.enTitleOptions);
    slide.addText(frAnnounce, { ...specificOptions.frContentOptions, fontSize: computeFontSize(frAnnounce)});
    slide.addText(enAnnounce, { ...specificOptions.enContentOptions, fontSize: computeFontSize(enAnnounce)});
  }
}

function computeFontSize(line) {
  if (line === undefined) return 36;
  let words = line.split(' ');
  let wordsNumber = words.length;
  switch (true) {
    case wordsNumber <= 10:
      return 50;
    case wordsNumber <= 17:
      return 50;
    case wordsNumber <= 20:
      return 47;
    case wordsNumber <= 25:
      return 47;


    //parfait
    case wordsNumber <= 30:
      return 36;
    case wordsNumber <= 35:
      return 38;

    case wordsNumber <= 40:
      return 32;
    case wordsNumber <= 45:
      return 32;
    case wordsNumber <= 47:
      return 30;
    case wordsNumber <= 50:
      return 29;

    //A ne plus toucher c'est parfait
    case wordsNumber <= 60:
      return 28;
    case wordsNumber <= 65:
      return 26;

    case wordsNumber <= 66:
      return 26;
    case wordsNumber <= 70:
      return 26;
    case wordsNumber <= 75:
      return 25;

    case wordsNumber <= 80:
      return 25;
    default:
      return 36
  }
}

function configureMasterSlide(pptx) {
  let slideMasterOptions = {
    title: 'MASTER',
    bkgd: { path: `${path.resolve()}/services/bg1.jpeg` }
  };

  let contentSlideMasterOptions = {
    title: 'MASTER_CONTENT',
    bkgd: { path: `${path.resolve()}/services/bg1.jpeg` },
    objects: [
      { 'placeholder': { options: { name: 'frTitle', type: 'body', x: 0.08, y: 0.12, w: 6.52, h: 0.75 } } },
      { 'placeholder': { options: { name: 'enTitle', type: 'body', x: 6.74, y: 0.12, w: 6.52, h: 0.75 } } },
      { 'placeholder': { options: { name: 'frContent', type: 'body', x: 0.08, y: 1.06, w: 6.52, h: 6.44 } } },
      { 'placeholder': { options: { name: 'enContent', type: 'body', x: 6.74, y: 1.06, w: 6.52, h: 6.44 } } }
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
    fontSize: 54,
    color: 'ffffff',
    align: 'center',
    fill: {
      type: 'solid',
      color: '000000',
      alpha: 40
    },
    isTextBox: true,
    shrinkText: true,
    rectRadius: 2,
    lineSpacing: 80
  }
};

const themOptions = {
  fontSize: 48,
  align: 'center',
  fill: {
    type: 'solid',
    color: '000000',
    alpha: 40
  },
  isTextBox: true,
  shrinkText: true,
  rectRadius: 2,
  lineSpacing: 130
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
    ...generalOptions.common
  },
  enContentOptions: {
    placeholder: 'enContent',
    ...generalOptions.common
  }
};


//officegen

function buildPPT() {
  let pptx = officegen('pptx');

  pptx.on('finalize', function (written) {
    console.log(
      'Finish to create a Microsoft PowerPoint document.'
    )
  })

  // Officegen calling this function to report errors:
  pptx.on('error', function (err) {
    console.log(err)
  })

}

function addSlide(text, options) {

}

function addText() {

}


exports.build = buildPPT;
exports.genPPT = buildpowerpoint;
