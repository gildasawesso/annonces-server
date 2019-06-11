let express = require('express');
let router = express.Router();

router.get('/', (req, res) => {
  res.json(Date.now());
});

module.exports = router;
