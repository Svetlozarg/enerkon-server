const mongoose = require("mongoose");

/*
  Project Schema
  _id: string
  title: string
  documents: [{
    fileName: string
    originalName: string
    type: string
    size: number
  }]
  favourite: boolean
  status: "Paid" | "Unpaid"
  createdAt: Date
  deletedAt: Date
*/

const projectSchema = new mongoose.Schema(
  {
    title: {
      type: String,
      required: [true, "Title field is required"],
    },
    documents: [
      {
        type: [Object],
        required: [true, "Documents field is required"],
      },
    ],
    favourite: {
      type: Boolean,
      default: false,
    },
    status: {
      type: String,
      enum: ["Paid", "Unpaid"],
      default: "Unpaid",
    },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Project", projectSchema);
