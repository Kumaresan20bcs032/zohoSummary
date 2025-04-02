require("dotenv").config();
const axios = require("axios");

// This function related to date format
const formatDate = (dateString) => {
    if (!dateString) {
        console.error("❌ Invalid Date Provided:", dateString);
        return null;
    }

    const date = new Date(dateString);
    const offset = -date.getTimezoneOffset();
    const offsetHours = String(Math.floor(Math.abs(offset) / 60)).padStart(2, "0");
    const offsetMinutes = String(Math.abs(offset) % 60).padStart(2, "0");
    const offsetSign = offset >= 0 ? "+" : "-";

    return `${date.toISOString().split(".")[0]}${offsetSign}${offsetHours}:${offsetMinutes}`;
};


module.exports = {
    formatDate
}
