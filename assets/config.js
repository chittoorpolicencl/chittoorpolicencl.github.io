/* ------------------------------------------------------------------
   SETTINGS — edit this file when stations or the repository change.
   ------------------------------------------------------------------ */
window.NCL_CONFIG = {
  /* Headings */
  unitName: "CHITTOOR DISTRICT POLICE",
  appTitle: "NEW CRIMINAL LAWS - 60 & 90 DAYS CASES",
  fromDate: "01-07-2024",            // shown in the abstract title
  credit: "Designed & Developed by Kola.Bapuji, IT Core Team, Chittoor",

  /* GitHub repository that hosts this site.
     Leave owner/repo blank to detect them from the address
     (https://OWNER.github.io/REPO/). Leave branch blank to use the
     repository's default branch. "folder" is only needed when the site
     is not at the top of the repository (for example "docs"). */
  github: { owner: "", repo: "", branch: "", folder: "" },

  /* Sub-divisions and their police stations, in the order they appear
     in the abstract (same as the SD sheet of the Excel workbook). */
  subDivisions: [
    { code: "CTR", title: "CHITTOOR - SUB DIVISION", totalLabel: "CHITTOOR SD TOTAL",
      stations: ["Chittoor I Town UPS", "Chittoor II Town UPS", "Chittoor Taluk PS",
                 "G.D.Nellore UPS", "Gudipala", "Irala", "Kallur", "Kanipakam",
                 "Narsingarayanipet(N.R.Pet)", "Penumur", "Puthalapattu UPS",
                 "Rompicherla", "Thavanampalle", "Yadamari"] },
    { code: "KPM", title: "KUPPAM SUB DIVISION", totalLabel: "KUPPAM SD TOTAL",
      stations: ["Gudipalli", "Kuppam UPS", "Rallabudugur", "Ramakuppam"] },
    { code: "PLMR", title: "PALAMANER SUB DIVISION", totalLabel: "PALAMANER SD TOTAL",
      stations: ["Baireddypalle", "Bangarupalem UPS", "Gangavaram UPS",
                 "Palamaner UPS", "Panjani", "V.Kota UPS"] },
    { code: "NGR", title: "NAGARI SUB DIVISION", totalLabel: "NAGARI SD TOTAL",
      stations: ["K. Nagar", "Nagari UPS", "Nindra", "Palasamudram", "S.R.Puram",
                 "Vedurukuppam", "Vijayapuram"] },
    { code: "OTHERS", title: "OTHERS", totalLabel: "OTHERS TOTAL",
      stations: ["CCS Chittoor UPS", "Chittoor Traffic UPS", "Mahila UPS, Chittoor"] }
  ],

  /* Rows of these stations are removed while importing (same as the macro). */
  excludeStations: ["Punganur UPS", "Sodam", "Somala", "Chowdepalli"]
};
