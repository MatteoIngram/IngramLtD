Ingram Ltd website
==================

Structure
  index.html        page markup
  css/site.css      styles
  js/main.js        builds the tables and chart from data/bdc2.js
  data/bdc2.js      index data (generated, do not edit by hand)
  export_data.py    Excel model -> data/bdc2.js
  docs/             put factsheet.pdf and methodology.pdf here
  img/              logo.svg and favicon.svg (replace with your own logo)

Monthly update
  1. Update and save BDC2_Index_Fund_Model.xlsx in Excel
  2. python export_data.py BDC2_Index_Fund_Model.xlsx
  3. Commit data/bdc2.js (and the new factsheet PDF in docs/)

Hosting: GitHub Pages, Settings > Pages > deploy from main / root.

To do before launch
  - real email address (index.html, search "hello@ingram.example")
  - factsheet.pdf and methodology.pdf in docs/
  - own logo in img/
