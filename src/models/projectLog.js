const mongoose = require("mongoose");

/*
  Project Log Schema
  _id: string
  id: string
  log: [{
    _id: string
    title: string
    action: string
    date: date
  }]
*/

const projectLogSchema = new mongoose.Schema({
  id: {
    type: String,
    required: [true, "Project ID is required"],
  },
  log: [
    {
      title: {
        type: String,
        required: [true, "Title field is required"],
      },
      action: {
        type: String,
        required: [true, "Action field is required"],
      },
      date: {
        type: Date,
        required: [true, "Date field is required"],
      },
    },
  ],
});

module.exports = mongoose.model("ProjectLog", projectLogSchema);
