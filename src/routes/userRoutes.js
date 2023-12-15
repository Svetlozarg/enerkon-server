const express = require("express");
const { currentUser } = require("../controllers/userController");
const validateToken = require("../middleware/validateTokenHandler");

const router = express.Router();

router.get("/current", validateToken, currentUser);

module.exports = router;
