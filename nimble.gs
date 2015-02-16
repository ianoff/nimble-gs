/**
 * Nimble
 * @constructor
 * @return {Object} The Nimble Object
 * @param {object}
 */
var Nimble = function(options) {
    // Make sure Nimble is being used as a constructor no matter what.
    if (!this || !(this instanceof Nimble)) {
        return new Nimble(options);
    }

    options = options || {};

    /**
    * The options object can contain a create object with has the following properties:
    * @example
    create: {
        name: "My New Sheet", //(optional)
        sheets: ["A","B","C"], //(optional)
        headers: ["Col1","Col2","Col3"],//(optional)
        clean: true//(optional, only if there are headers)
    }
    */
    if (options.create) {
        if (options.create.name) {
            this.ss = SpreadsheetApp.create(options.create.name);
        } else {
            this.ss = SpreadsheetApp.create("New Spreadsheet");
        }
        this.url = this.ss.getUrl();

        function createHeaders(ss) {
            if (options.create.headers) {
                var wrap = [];
                wrap[0] = options.create.headers;
                ss.getRange(1, 1, 1, options.create.headers.length).setValues(wrap);
            }
        }

        if (options.create.sheets && options.create.sheets instanceof Array) {
            for (var i = 0; i < options.create.sheets.length; i++) {
                var sheetName = options.create.sheets[i];
                this.ss.insertSheet(i + 1);
                this.ss.renameActiveSheet(sheetName);
                this.sheet = this.ss.getActiveSheet();
                createHeaders(this.sheet);
                if (options.create.headers && options.create.clean) this.cleanEmptyCols();
            }
        } else {
            this.ss.insertSheet(1);
            this.sheet = this.ss.getActiveSheet();
            createHeaders(this.ss.getActiveSheet());
            if (options.create.headers && options.create.clean) this.cleanEmptyCols();
        }
        this.ss.deleteSheet(this.ss.getSheetByName("Sheet1"));
        this.sheet = this.ss.getSheets()[0];
    }

    //NOTE: Just found out the "active" related functions only work with bound scripts
    // Hafta rework
    // See: https://developers.google.com/apps-script/guides/bound 

    /* Setup of Spreadsheet
    url
    default: the active apreadsheet attached to the project
    optional: pass in the url as a string
    */
    if (!options.create && typeof(options.url) == "string") {

      if (/^http/.test(options.url)) {
        this.ss = SpreadsheetApp.openByUrl(options.url);
      } else {
        this.ss = SpreadsheetApp.openById(options.url);
      }
        
    } else {
        this.ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    /* Setup of this.sheet
    this.sheet
    default: the active this.sheet in the active spreadsheet
    optional: pass in a number (zero-indexed position) or string (name) of a this.sheet
     */
    if (!options.create && typeof(options.sheet) == "number") {
        this.sheet = this.ss.getSheets()[options.sheet];
    } else if (typeof(options.sheet) == "string") {
        this.sheet = this.ss.getSheetByName(options.sheet);
    } else {
        this.sheet = this.ss.getActiveSheet();
    }
    return this;
};

Nimble.prototype = {
    /**
     * @function getHeaders
     * @return {Object} An array of the column names in the order they appear
     */
    getHeaders: function() {
        var colNamesArray = this.sheet.getRange(1, 1, 1, this.sheet.getMaxColumns()).getValues();
        return colNamesArray[0];
    },
    /**
     * @return {Object} gets row position of a property. If row doesn't exist, returns "0"
     */
    getCol: function(property) {
        var cols = this.getHeaders(this.sheet);
        return cols.indexOf(property) + 1;
    },

    getColRange: function(property, startRow, endRow) {
        var col = this.getCol(property);
        if (col != 0) {
            if (endRow == undefined) {
                last = this.sheet.getLastRow();
            } else {
                last = endRow;
            }
            var range = this.sheet.getRange(startRow, col, last, 1);
            return range;
        }
    },

    getColData: function(property, startRow, endRow) {
        var range = this.getColRange(property, startRow, endRow);
        return range.getValues();
    },

    /*Gets the last cell in a column*/

    getLastCell: function(property) {
        var col = this.getCol(property);
        var lastCell = this.sheet.getRange(this.sheet.getLastRow(), col);
        return lastCell;
    },

    getMaxCell: function(property) {
        var col = this.getCol(property);
        var lastCell = this.sheet.getRange(this.sheet.getMaxRows(), col);
        return lastCell;
    },

    /*Gets the last cell with data in a column*/

    emptyStart: function(property) {
        var colData = this.getColData(property, 2);
     
        for (var i = 0; i < colData.length; ++i) {
          info = colData[i][0];
            if (info == "") {
                return i + 2;
            }
        }
    },

    /*Adds a new column after the last existing column*/

    addCol: function(property) {
        this.sheet.insertColumnAfter(this.sheet.getMaxColumns());
        newCol = this.sheet.getMaxColumns();
        this.sheet.getRange(1, newCol).setValue(property);
    },

    /*Deletes a column by name*/

    deleteCol: function(property) {
        var propCol = this.getCol(property);
        if (propCol !== 0) {
            this.sheet.deleteColumn(propCol);
        }
    },

    /*Deletes several columns by name*/

    deleteCols: function() {
        for (var i = 0; i < arguments.length; ++i) {
            var property = arguments[i];
            this.deleteCol(property);
        }
    },

    /*Renames a column header*/

    renameCol: function(property, newName) {
        propCol = this.getCol(property);
        if (propCol !== 0) {
            this.sheet.getRange(1, this.getCol(property)).setValue(newName);
        }
    },

    cleanEmptyRows: function() {
        content = this.sheet.getLastRow();
        all = this.sheet.getMaxRows();
        if (all - content > 0) {
            this.sheet.deleteRows(this.sheet.getLastRow() + 1, all - content);
        }
    },
    cleanEmptyCols: function() {
        content = this.sheet.getLastColumn();
        all = this.sheet.getMaxColumns();
        if (all - content > 0) {
            this.sheet.deleteColumns(this.sheet.getLastColumn() + 1, all - content);
        }
    },
    cleanEmpty: function() {
        this.cleanEmptyCols();
        this.cleanEmptyRows();
    },

    makeJSON: function() {
        var rowsData = this.getRowsData(false);
        return Utilities.jsonStringify(rowsData);
    },
    makeGeoJSON: function() {
        var rowsData = this.getRowsData(true);
        return Utilities.jsonStringify(rowsData);
    },
    serveJSON: function() {
        var json = this.makeJSON();
        return ContentService.createTextOutput(json)
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    },
    serveJSONP: function(callback) {
        var json = this.makeJSON();
        return ContentService.createTextOutput(
            callback + '(' + json + ');')
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    },

    /* Based on https://developers.google.com/apps-script/guides/sheets#reading
  and 
  http://blog.pamelafox.org/2013/06/exporting-google-spreadsheet-as-json.html*/

    getRowsData: function(geo) {
        var lastCol = this.sheet.getMaxColumns(),
            lastRow = this.sheet.getMaxRows();

        var headers = this.getHeaders(this.sheet);
        var dataRange = this.sheet.getRange(2, 1, lastRow, lastCol);
        if (!geo) {
            var objects = this.getObjects(dataRange.getValues(), this.normalizeHeaders(headers));
        } else {
            objects = this.getGeoObjects(dataRange.getValues(), this.normalizeHeaders(headers));
        }
        return objects;

    },

    getObjects: function(data, keys) {
        var objects = [];
        for (var i = 0; i < data.length; ++i) {
            var object = {};
            var hasData = false;
            for (var j = 0; j < data[i].length; ++j) {
                var cellData = data[i][j];
                if (this.isCellEmpty(cellData)) {
                    continue;
                }
                object[keys[j]] = cellData;
                hasData = true;
            }
            if (hasData) {
                objects.push(object);
            }
        }
        return objects;
    },

    getGeoObjects: function(data, keys) {
        var latCol = (this.getCol("lat") - 1);
        var lngCol = (this.getCol("lng") - 1);
        var objects = [{
            "type": "FeatureCollection",
            "features": []
        }];
        for (var i = 0; i < data.length; ++i) {
            var object = {
                "type": "Feature",
                "properties": {},
                "geometry": {
                    "type": "Point",
                    "coordinates": []
                }
            };
            var hasData = false;
            for (var j = 0; j < data[i].length; ++j) {

                var cellData = data[i][j];
                if (this.isCellEmpty(cellData)) {
                    continue;
                }
                if (j == latCol) {
                    object.geometry.coordinates.push(cellData)
                } else if (j == lngCol) {
                    object.geometry.coordinates.unshift(cellData)
                } else {
                    object.properties[keys[j]] = cellData;
                }
                hasData = true;
            }
            if (hasData) {
                objects[0].features.push(object);
            }
        }
        return objects;
    },

    // Returns an Array of normalized Strings.
    // Arguments:
    //   - headers: Array of Strings to normalize

    normalizeHeaders: function(headers) {
        var keys = [];
        for (var i = 0; i < headers.length; ++i) {
            var key = this.normalizeHeader(headers[i]);
            if (key.length > 0) {
                keys.push(key);
            }
        }
        return keys;
    },

    // Normalizes a string, by removing all alphanumeric characters and using mixed case
    // to separate words. The output will always start with a lower case letter.
    // This function is designed to produce JavaScript object property names.
    // Arguments:
    //   - header: string to normalize
    // Examples:
    //   "First Name" -> "firstName"
    //   "Market Cap (millions) -> "marketCapMillions
    //   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"

    normalizeHeader: function(header) {
        var key = "";
        var upperCase = false;
        for (var i = 0; i < header.length; ++i) {
            var letter = header[i];
            if (letter == " " && key.length > 0) {
                upperCase = true;
                continue;
            }
            if (!this.isAlnum(letter)) {
                continue;
            }
            if (key.length == 0 && this.isDigit(letter)) {
                continue; // first character must be a letter
            }
            if (upperCase) {
                upperCase = false;
                key += letter.toUpperCase();
            } else {
                key += letter.toLowerCase();
            }
        }
        return key;
    },

    // Returns true if the cell where cellData was read from is empty.
    // Arguments:
    //   - cellData: string

    isCellEmpty: function(cellData) {
        return typeof(cellData) == "string" && cellData == "";
    },

    // Returns true if the character char is alphabetical, false otherwise.

    isAlnum: function(char) {
        return char >= 'A' && char <= 'Z' ||
            char >= 'a' && char <= 'z' ||
            this.isDigit(char);
    },

    // Returns true if the character char is a digit, false otherwise.

    isDigit: function(char) {
        return char >= '0' && char <= '9';
    },

    // Given a JavaScript 2d Array, this function returns the transposed table.
    // Arguments:
    //   - data: JavaScript 2d Array
    // Returns a JavaScript 2d Array
    // Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].

    arrayTranspose: function(data) {
        if (data.length == 0 || data[0].length == 0) {
            return null;
        }

        var ret = [];
        for (var i = 0; i < data[0].length; ++i) {
            ret.push([]);
        }

        for (var i = 0; i < data.length; ++i) {
            for (var j = 0; j < data[i].length; ++j) {
                ret[j][i] = data[i][j];
            }
        }

        return ret;
    },

    /*Loops through a range of values and replaces with a new value after being altered by a function*/

    manipData: function(property, functionName, startRow, endRow) {
        var arr = this.getColData(property, startRow, endRow),
            newArr = new Array(arr.length);
        for (var i = 0; i < arr.length; ++i) {
            var item = arr[i],
                newValue = functionName(item[0]);
            if (!newValue instanceof Array) {
                newArr[i] = newValue;
            } else {
                var shell = [];
                shell[0] = newValue;
                newArr[i] = shell;
            }
        }
        this.sheet.getRange(startRow, this.getCol(property), arr.length).setValues(newArr);
    },

    /*truncates to a specified letter*/
    truncateValue: function(property, numLetters, startRow, endRow) {
        function trunc(str) {
            return str.slice(0, numLetters);
        }
        this.manipData(property, trunc, startRow, endRow);
    },

    /*sets string to lowercase*/
    lowercase: function(property, startRow, endRow) {
        function low(str) {
            return str.toLowerCase();
        }
        this.manipData(property, low, startRow, endRow);
    },

    /** Geolocator **/
    /**
    locInput: column name or array of column names
    adj: boolean. Should an adjusted lat & lng be calculated?
    **/
    geocode: function(locInput, adj) {
        //Create new columns if needed
        if (this.getCol("lat") == 0) {
            this.addCol("lat");
            this.addCol("lng");
        }

        if (adj) {
            if (this.getCol("adj_lat") == 0) {
                this.addCol("adj_lat");
                this.addCol("adj_lng");
            }
        }

        //save cols as var
        var latCol = this.getCol("lat"),
            lngCol = this.getCol("lng"),
            startingRow = this.emptyStart("lat"),
            data = {},
            lat,
            lng,
            adjlat,
            adjlng;

        if (locInput instanceof Array) {
            list = this.getColData(locInput[0], startingRow).length;
        } else {
            list = this.getColData(locInput, startingRow).length;
        }


        if (adj) {
            var adjlatCol = this.getCol("adj_lat"),
                adjlngCol = this.getCol("adj_lng"),
                angle = 1, // starting angle, in radians
                lng_radius = 1, // degrees of longitude separation
                lat_to_lng = 111.23 / 71.7, // lat to long proportion in Warsaw
                step = 200 * Math.PI / list,
                lat_radius = lng_radius / lat_to_lng;
        }

        for (var i = 0; i < list; i++) {
            geocodeThis = "";

            cell = (i + startingRow);

            for (var j = 0; j < locInput.length; j++) {
                geocodeThis = j > 0 ? geocodeThis + ", " + this.sheet.getRange(cell, this.getCol(locInput[j])).getValue() : geocodeThis + this.sheet.getRange(cell, this.getCol(locInput[j])).getValue();
            }

            geocodeThis = geocodeThis + "";

            if (/national/i.test(geocodeThis.toLowerCase())) {
                //Washington DC
                lat = 38.906881;
                lng = -77.038193;
            } else {
                var response = Maps.newGeocoder().geocode(geocodeThis);
                lat = response.results[0].geometry.location.lat; //latitude as returned by Google Maps
                lng = response.results[0].geometry.location.lng; //longitude as returned by Google Maps
            }

            this.sheet.getRange(cell, latCol).setValue(lat); //Set the actual Lat
            this.sheet.getRange(cell, lngCol).setValue(lng); //Set the actual Lng

            if (adj) {
                adjlat = lat + (Math.cos(angle) * lng_radius);
                adjlng = lng + (Math.sin(angle) * lat_radius);
                angle += step;
                this.sheet.getRange(cell, adjlatCol).setValue(adjlat); //Set the adj Lat
                this.sheet.getRange(cell, adjlngCol).setValue(adjlng); //Set the adj Lng
            }

            //Timeout so we don't blow past the time-based geocoding limitations
            Utilities.sleep(1000);

        }
    },
    /** Geolocator **/
    /**
    locInput: array of column names
    adj: boolean. Should an adjusted lat & lng be calculated?
    **/
    geocodeBatch: function(locInput, adj) {
        //Create new columns if needed
        if (this.getCol("lat") == 0) {
            this.addCol("lat");
            this.addCol("lng");
        }

        if (adj) {
            if (this.getCol("adj_lat") == 0) {
                this.addCol("adj_lat");
                this.addCol("adj_lng");
            }
        }

        //save cols as var
        var latCol = this.getCol("lat"),
            lngCol = this.getCol("lng"),
            startingRow = this.emptyStart("lat"),
            data = {},
            lat,
            lng,
            adjlat,
            adjlng,
            latArr = [],
            lngArr = [],
            adjLatArr = [],
            adjLngArr = [];
      
      if (startingRow == "undefined") return;
      
        var maxSize = 350;

        var list = this.getColData(locInput[0], startingRow).length < maxSize ? this.getColData(locInput[0], startingRow).length : maxSize;


        if (adj) {
            var adjlatCol = this.getCol("adj_lat"),
                adjlngCol = this.getCol("adj_lng"),
                angle = 1, // starting angle, in radians
                lng_radius = 1, // degrees of longitude separation
                lat_to_lng = 111.23 / 71.7, // lat to long proportion in Warsaw
                step = 200 * Math.PI / list,
                lat_radius = lng_radius / lat_to_lng;
        }




        for (var i = 0; i < list; i++) {
            geocodeThis = "";

            cell = (i + startingRow);
          
          

            for (var j = 0; j < locInput.length; j++) {
                geocodeThis = j > 0 ? geocodeThis + ", " + this.sheet.getRange(cell, this.getCol(locInput[j])).getValue() : geocodeThis + this.sheet.getRange(cell, this.getCol(locInput[j])).getValue();
            }

            geocodeThis = geocodeThis + "";

            if (/national/i.test(geocodeThis.toLowerCase())) {
                //Washington DC
                lat = 38.906881;
                lng = -77.038193;
            } else {
                var response = Maps.newGeocoder().geocode(geocodeThis);
              
              if (response.status != "ZERO_RESULTS") {
                lat = response.results[0].geometry.location.lat; //latitude as returned by Google Maps
                lng = response.results[0].geometry.location.lng; //longitude as returned by Google Maps
              } 
              
            }

            //Lat
            var l = []
            l.push(lat);
            latArr.push(l);

            //Lng
            var m = []
            m.push(lng);
            lngArr.push(m);



          

            if (adj) {
                adjlat = lat + (Math.cos(angle) * lng_radius);
                adjlng = lng + (Math.sin(angle) * lat_radius);
                angle += step;
                //Lat
                var o = []
                o.push(adjlat);
                adjLatArr.push(o);

                //Lng
                var p = []
                p.push(adjlng);
                adjLngArr.push(p);
            }

            //Timeout so we don't blow past the time-based geocoding limitations
            Utilities.sleep(300);

        }
      
      this.sheet.getRange(startingRow, latCol, list).setValues(latArr);
      this.sheet.getRange(startingRow, lngCol, list).setValues(lngArr);
      this.sheet.getRange(startingRow, adjlatCol, list).setValues(adjLatArr);
      this.sheet.getRange(startingRow, adjlngCol, list).setValues(adjLngArr);
    }
};
