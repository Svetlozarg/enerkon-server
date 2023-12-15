const mongoose = require("mongoose");

const userSchema = mongoose.Schema(
  {
    username: {
      type: String,
      required: [true, "Username field is required"],
    },
    email: {
      type: String,
      required: [true, "Email address field is required"],
      unique: [true, "Email address already taken"],
    },
    password: {
      type: String,
      required: [true, "Password field is required"],
    },
  },
  {
    timestamps: true,
  }
);

module.exports = mongoose.model("User", userSchema);
