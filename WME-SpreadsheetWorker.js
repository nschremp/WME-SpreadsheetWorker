  // ==UserScript==
  // @name             Spreadsheet Worker
  // @namespace        https://greasyfork.org/en/users/77740-nathan-fastestbeef-fastestbeef
  // @version          2020.04.29.bootstrap
  // @description      makes working spreadsheet projects easier
  // @author           FastestBeef
  // @include          https://www.waze.com/editor*
  // @include          https://www.waze.com/*/editor*
  // @include          https://beta.waze.com/editor*
  // @include          https://beta.waze.com/*/editor*
  // @exclude          https://www.waze.com/*user/editor*
  // @grant            none
  // @require          https://greasyfork.org/scripts/24851-wazewrap/code/WazeWrap.js
  // ==/UserScript==

  /* global W */
  /* ecmaVersion 2017 */
  /* global $ */
  /* global WazeWrap */
  /* global OpenLayers */

  (function() {
      'use strict';

      const VERSION = GM_info.script.version;
      const SCRIPT_NAME = GM_info.script.name;
      const UPDATE_NOTES = `
  <p>
    <ul>
      <li>Bug fix: State filter did not work for states with two words in the name</li>
    </ul>
  </p>`;

      const CAMPAIGN_SHEET_ID = "1vy_28mDW8CDIUUXbUG0XZgCNX7iJ6sdPEtbCvQvXrLs";

      const STATES = [
        { "name": "Alabama", "abbreviation": "AL" },
        { "name": "Alaska", "abbreviation": "AK" },
  //      { "name": "American Samoa", "abbreviation": "AS" },
        { "name": "Arizona", "abbreviation": "AZ" },
        { "name": "Arkansas", "abbreviation": "AR" },
        { "name": "California", "abbreviation": "CA" },
        { "name": "Colorado", "abbreviation": "CO" },
        { "name": "Connecticut", "abbreviation": "CT" },
        { "name": "Delaware", "abbreviation": "DE" },
        { "name": "District Of Columbia", "abbreviation": "DC" },
  //      { "name": "Federated States Of Micronesia", "abbreviation": "FM" },
        { "name": "Florida", "abbreviation": "FL" },
        { "name": "Georgia", "abbreviation": "GA" },
  //      { "name": "Guam Gu", "abbreviation": "GU" },
        { "name": "Hawaii", "abbreviation": "HI" },
        { "name": "Idaho", "abbreviation": "ID" },
        { "name": "Illinois", "abbreviation": "IL" },
        { "name": "Indiana", "abbreviation": "IN" },
        { "name": "Iowa", "abbreviation": "IA" },
        { "name": "Kansas", "abbreviation": "KS" },
        { "name": "Kentucky", "abbreviation": "KY" },
        { "name": "Louisiana", "abbreviation": "LA" },
        { "name": "Maine", "abbreviation": "ME" },
  //      { "name": "Marshall Islands", "abbreviation": "MH" },
        { "name": "Maryland", "abbreviation": "MD" },
        { "name": "Massachusetts", "abbreviation": "MA" },
        { "name": "Michigan", "abbreviation": "MI" },
        { "name": "Minnesota", "abbreviation": "MN" },
        { "name": "Mississippi", "abbreviation": "MS" },
        { "name": "Missouri", "abbreviation": "MO" },
        { "name": "Montana", "abbreviation": "MT" },
        { "name": "Nebraska", "abbreviation": "NE" },
        { "name": "Nevada", "abbreviation": "NV" },
        { "name": "New Hampshire", "abbreviation": "NH" },
        { "name": "New Jersey", "abbreviation": "NJ" },
        { "name": "New Mexico", "abbreviation": "NM" },
        { "name": "New York", "abbreviation": "NY" },
        { "name": "North Carolina", "abbreviation": "NC" },
        { "name": "North Dakota", "abbreviation": "ND" },
  //    { "name": "Northern Mariana Islands", "abbreviation": "MP" },
        { "name": "Ohio", "abbreviation": "OH" },
        { "name": "Oklahoma", "abbreviation": "OK" },
        { "name": "Oregon", "abbreviation": "OR" },
  //    { "name": "Palau", "abbreviation": "PW" },
        { "name": "Pennsylvania", "abbreviation": "PA" },
  //    { "name": "Puerto Rico", "abbreviation": "PR" },
        { "name": "Rhode Island", "abbreviation": "RI" },
        { "name": "South Carolina", "abbreviation": "SC" },
        { "name": "South Dakota", "abbreviation": "SD" },
        { "name": "Tennessee", "abbreviation": "TN" },
        { "name": "Texas", "abbreviation": "TX" },
        { "name": "Utah", "abbreviation": "UT" },
        { "name": "Vermont", "abbreviation": "VT" },
  //    { "name": "Virgin Islands", "abbreviation": "VI" },
        { "name": "Virginia", "abbreviation": "VA" },
        { "name": "Washington", "abbreviation": "WA" },
        { "name": "West Virginia", "abbreviation": "WV" },
        { "name": "Wisconsin", "abbreviation": "WI" },
        { "name": "Wyoming", "abbreviation": "WY" }
      ];

      let settings = {};
      let sheetData = {};
      let tab = {};
      let campaigns = [];

      let pgMaxX = 0;
      let pgMinX = 0;
      let pgMaxY = 0;
      let pgMinY = 0;

      function getCampaignData() {
          let apiKey = localStorage.getItem('SW_API_KEY')

          if( apiKey === '' ) {
              WazeWrap.Alerts.error('Spreadsheet Worker', 'You must set an API key before using this script');
              return;
          }

          let url = "https://sheets.googleapis.com/v4/spreadsheets/"+CAMPAIGN_SHEET_ID+"/values/Sheet1!A1:J500?key="+apiKey;
          console.log("SW: getting sheet info ("+url+")");

          $.ajax({
              url: url,
              success: function(data){
                $('#swCampaignSelect')
                  .empty()
                  .append('<option selected="selected" value="">Select</option>');
                  data.values.shift(); // Kill header row
                  data.values.forEach(function(item, i){
                      let campaignRow = {campaignName:item[0],
                                         spreadsheetId:item[1],
                                         sheetName:item[2],
                                         maxRows:item[3],
                                         lonCol:item[4],
                                         latCol:item[5],
                                         stateCol:item[6],
                                         completeCol:item[7],
                                         betaOnly:(item[8]==='TRUE'),
                                         active:(item[9]==='TRUE')};
                      campaigns.push(campaignRow);
                      if(campaignRow.active){$('#swCampaignSelect').append(new Option(item[0], i))};
                  });
              },
              dataType: 'JSON'
          });
      }

      function getAllRowData() {
          let campaignRow=document.getElementById('swCampaignSelect').value;
          let spreadsheetId=campaigns[campaignRow].spreadsheetId;
          let sheetName=campaigns[campaignRow].sheetName;
          let maxRows=campaigns[campaignRow].maxRows;
          let apiKey = localStorage.getItem('SW_API_KEY')

          if( apiKey === '' ) {
              WazeWrap.Alerts.error('Spreadsheet Worker', 'You must set an API key before using this script');
              return;
          }

          let url = encodeURI("https://sheets.googleapis.com/v4/spreadsheets/"+spreadsheetId+"/values/"+sheetName+"!A1:J"+maxRows+"?key="+apiKey);
          console.log("SW: getting sheet info ("+url+")");

          $.ajax({
              url: url,
              success: function(data){
                  sheetData = data;
              },
              dataType: 'JSON'
          });

          document.getElementById('swCurRow').value = 1;
      }

      function getPrev() {
          $('#swCurRow').val( function(i, oldval) {
              return parseInt(oldval, 10) - 2;
          });
          getNext();
      }

      function stateFilterPass(state) {
          return state.toLowerCase() === $('#swStateFilter').val().toLowerCase() ||
                 '' === $('#swStateFilter').val()
      }

      function polygonPass(xCoord, yCoord) {
        if( xCoord < pgMinX || xCoord > pgMaxX || yCoord < pgMinY || yCoord > pgMaxY) {
          return false;
        }
        return true;
      }

      function getNext() {
          let campaignRow = $('#swCampaignSelect').val();

          if (campaignRow === '') {
              WazeWrap.Alerts.error('Spreadsheet Worker', 'You must select a campaign first.');
          }
          let completeCol = campaigns[campaignRow].completeCol.toUpperCase().charCodeAt(0) - 65;
          let stateCol = campaigns[campaignRow].stateCol.toUpperCase().charCodeAt(0) - 65;
          let lonCol = campaigns[campaignRow].lonCol.toUpperCase().charCodeAt(0) - 65;
          let latCol = campaigns[campaignRow].latCol.toUpperCase().charCodeAt(0) - 65;

          let currentRow = parseInt($('#swCurRow').val(), 10);

          while(typeof sheetData.values[currentRow] !== "undefined" && currentRow < 500000 ) {
              let lon = sheetData.values[currentRow][lonCol];
              let lat = sheetData.values[currentRow][latCol];

              if( (typeof sheetData.values[currentRow][completeCol] === "undefined" ||
                   sheetData.values[currentRow][completeCol] === "") &&
                  stateFilterPass(sheetData.values[currentRow][stateCol]) &&
                  polygonPass(lon, lat)
                ) {
                  console.log("SW: Row="+(currentRow+1)+" Lon="+lon+" Lat="+lat);

                  var location = OpenLayers.Layer.SphericalMercator.forwardMercator(parseFloat(lon), parseFloat(lat));

                  //W.map.getOLMap().zoomTo(9);
                  W.map.setCenter(location);
                  document.getElementById('swCurRow').value = currentRow+1;
                  return;
              }
              currentRow++;
          }
          WazeWrap.Alerts.info("Spreadsheet Worker", "No more rows found.");
      }

      function updateAPIKey() {
          let apiKey = $('#swAPIKey').val();
          if(apiKey){
              localStorage.setItem('SW_API_KEY', apiKey);
              getCampaignData();
          }
          else {
              WazeWrap.Alerts.error('Spreadsheet Worker', 'No API Key set.');
          }
      }

      function bootstrap(tries = 1) {
          if (W &&
              W.map &&
              W.model &&
              W.loginManager.user &&
              $ && WazeWrap.Ready) {
              init();
          }
          else if (tries < 1000) {
              setTimeout(function () {bootstrap(tries++);}, 200);
          }
      }

      bootstrap();

      function init(){
          console.log("SW: Spreadsheet Worker Initializing.");
          tab = new WazeWrap.Interface.Tab("SW", tabHTML(),
                                           function (){
              $("#swGetNextBtn").click(()=>{getNext()});
              STATES.forEach(function(item, i){
                  $('#swStateFilter').append("<option value='"+item.name+"'>"+item.name+"</option>");
              });
              $("#swStateFilter").change(()=>{document.getElementById('swCurRow').value = 1;});
              $("#swCampaignSelect").change(()=>{getAllRowData()});
              $("#swAPIKeyUpdate").click(function(){updateAPIKey();});
              $("#refreshCampaign").click(function(){getCampaignData();});
              //$("#swTab").tabs();
          });

          if (!localStorage.getItem('SW_API_KEY')) {
              localStorage.setItem('SW_API_KEY', '');
          }

          getCampaignData();

          new WazeWrap.Interface.Shortcut("Next Row", "Get next row from spreadsheet worker script", "wmessw", "WME Spreadsheet Worker", "N", function(){getNext();}, null).add();
          new WazeWrap.Interface.Shortcut("Prev Row", "Get previous row from spreadsheet worker script", "wmessw", "WME Spreadsheet Worker", "S+N", function(){getPrev();}, null).add();

          WazeWrap.Interface.ShowScriptUpdate(SCRIPT_NAME, VERSION, UPDATE_NOTES, "https://greasyfork.org/en/scripts/401655-spreadsheet-worker", "https://www.waze.com/forum/viewtopic.php?f=819&t=301076");

          console.log("SW: Spreadsheet Worker Initialized.");
      }

      function tabHTML(){
          return `
  <div id='swTab' class="">
    <ul class="nav nav-tabs" id="myTab" role="tablist">
      <li class="nav-item">
        <a class="nav-link active" id="home-tab" data-toggle="tab" href="#swHome" role="tab" aria-controls="swHome" aria-selected="true">Spreadsheet Worker</a>
      </li>
      <li class="nav-item">
        <a class="nav-link" id="settings-tab" data-toggle="tab" href="#swSettings" role="tab" aria-controls="settings" aria-selected="false">Settings</a>
      </li>
    </ul>
    <div class="tab-content">
      <div class="tab-pane active" id="swHome" role="tabpanel" aria-labelledby="swHome-tab">
        <div class='form-row'>
          <div class='form-group col-md-10'>
            <label for='swCampaignSelect'>Campaign</label>
            <select id='swCampaignSelect' class="form-control">
              <option value=''>Select</option>
            </select>
          </div>
          <div class='form-group col-md-2'>
            <button class="btn btn-outline-secondary" id="refreshCampaign">
              <span class="w-icon w-icon-refresh">refresh</button>
            </button>
          </div>
        </div>
      </div>
      <div class="tab-pane" id="swSettings" role="tabpanel" aria-labelledby="swSettings-tab">
        <p>2</p>
      </div>
    </div>
  </div>`;
  let dummy = `
    <div class="">
      <div id="swHome" class="">

        <div style='display: block' >
          <label for='swStateFilter'>Filter State</label>
          <select id='swStateFilter'>
            <option value=''>None</option>
          </select>
        </div>
        <div>
          <input type="file" id="swFilterPolygon" accept=".wkt">
        </div>
        <div>
          <button class="btn btn-primary" id="swFilterPolygonImport" title="Import">Import</button>
        </div>
        <div style='display: block' >
          <button id='swGetNextBtn'>Next</button>
          <label for='swCurRow'>Current Row</label>
          <input id='swCurRow' size=10 value=1 />
        </div>
      </div>
      <div id="swSettings" class="">
        <div style='' >
          <label for="swAPIKey">API Key</label>
          <input class="form-control" id="swAPIKey">
          <button type="button" class="btn btn-primary" id="swAPIKeyUpdate">Update API key</button>
        </div>
        <div>
          <ol>
            <li>Go to <a href='https://console.cloud.google.com/projectselector2/apis/credentials'>Google Cloud Console</a></li>
            <li>Select create a project</li>
            <li>Give it any name you want and click create</li>
            <li>Click create credentials -> API Key</li>
            <li>Click Dashboard on the left</li>
            <li>Click enable APIs and Services</li>
            <li>Find Google Sheets API and click it</li>
            <li>Click Enable.</li>
            <li>Click the back arrow on the top left twice.</li>
            <li>Click 'Credentials' on the left side</li
            <li>Copy the generated key, paste it into the above box, and click 'Update API Key'</li>
            <li>Done. You should be able to use the script. It may take a few minutes for the changes to take effect</li>
            <li>If the campaign select box is not populating, try refreshing the list</li>
          </ol>
        </div>
      </div>
    </div>
  </div>`;
      }
  })();
