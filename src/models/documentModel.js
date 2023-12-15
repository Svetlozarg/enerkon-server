const mongoose = require("mongoose");

/*
  Document Schema
  _id: string
  title: string
  documents: {
    fileName: string
    originalName: string
    size: number
  }
  project: string
  type: string
  status: "In process" | "Canceled" | "Finished"
  createdAt: Date
  deletedAt: Date
*/

const documentSchema = new mongoose.Schema(
  {
    title: {
      type: String,
      required: [true, "Title field is required"],
    },
    document: {
      type: Object,
      required: [true, "Document field is required"],
    },
    project: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Project",
      required: [true, "Project is required"],
    },
    type: {
      type: String,
      required: [true, "Type is required"],
    },
    status: {
      type: String,
      enum: ["In process", "Canceled", "Finished"],
      default: "In process",
    },
    default: {
      type: Boolean,
      default: false,
    },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Document", documentSchema);
