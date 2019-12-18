/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'New SSC Sheet', functionName: 'createCategorySheet'}
  ];
  spreadsheet.addMenu('Select SSC', menuItems);
}

function doPost(e){
  var value = {};
  var parameter = e.parameter;
  Logger.log(parameter);
  if(!parameter.hasOwnProperty('format')){
    delete parameter['format'];
  }
  
  if(!parameter.hasOwnProperty('functionName')){
    return ContentService.createTextOutput(JSON.stringify({errors:["Invalid Request"]})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheetName = parameter['functionName'];
  delete parameter['functionName'];
  switch(sheetName){
    case 'createSheet':
      value['sheetName'] = createSheet(parameter);
      break;
    case 'calculatePrice':
      value['price'] = calculatePrice(parameter);
      break;
    default:
      value['price'] = calculatePrice(parameter);
  } 
  
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); 
}

function createSheet(e) {
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  var sub_sub_category_id = e.sub_sub_category_id;
  //var sub_sub_category_id = 52;
  
  var response = getPSANDSS(sub_sub_category_id);
  //var response = rsc_corrugated_box();
  var sheetName = response.name;
  create_sheet(sheetName);
  var sheet = ss.getSheetByName(sheetName);
  sheet.activate();
  // var sheet = ss.insertSheet(sheetName, ss.getNumSheets());
  
  sheet.autoResizeColumns(1, 6);
  sheet.getRange(1, 1).setValue("Per Unit Price");
  sheet.getRange(2, 1).setValue("Other Outputs");
  sheet.getRange(3,1,1,2).setValues([["Specification*","Input Values*"]]);
  sheet.getRange(3,3,1,1).setValues([["Sample Value*"]]);
  sheet.getRange(3,1,1,3).setBackground("orange");
  populate_specification(sheet, response);
  createLayout(sheet);
  protect_sheet(sheet);
  
  return sheet.getName();
 }

function calculatePrice(e){
  var parameter = e;
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
 
  var sheet = ss.getSheetByName(parameter.sheetName); 
  sheet.activate();

  delete parameter['sheetName'];
  var properties = parameter;
  
  var last_non_empty_row = sheet.getRange("A4:A").getValues().filter(String).length;
  var specification_list = sheet.getRange(4, 1, last_non_empty_row, 1).getValues();
 
  for (var key in properties) {
    if (properties.hasOwnProperty(key)) {
      Logger.log(key);
      var row = get_product_specification_cell(specification_list, key);
       sheet.getRange(row, 2).setValue(properties[key]);
     }
   }
 
  return sheet.getRange(1, 2).getValue();
}

function test(){
  var e = {
    'sheetName': "Corrugated Boxes sheet",
    'Material': 'Siddhesh',
  };
  
  calculatePrice(e); 
}


//layout work and app integration check I have to do

function createLayout(sheet){
  sheet.activate();
  
  sheet.getRange("A1:F").setWrap(true);
  var last_non_empty_row = sheet.getRange("A4:A").getValues().filter(String).length;
  
  // sheet.getRange(4, 1, last_non_empty_row + 3, 1).setBackground("#dddddd");
  // var range = sheet.getRange(1, 1, last_non_empty_row + 3, 6);
  
  //var range = sheet.getRange(3, 1, 1, 3);
  //var protection = range.protect().setDescription('protected range');
  //create_protect_ranges(2, 0, 2, 3, sheet_id);
  
  // var me = Session.getActiveUser();
  // protection.addEditor('service-account@spreadsheet-rajabose.iam.gserviceaccount.com');
  // protection.addEditors(['service-account@spreadsheet-rajabose.iam.gserviceaccount.com']);
  
  // protection.removeEditors(protection.getEditors());
  // if (protection.canDomainEdit()) {
   //  protection.setDomainEdit(false);
 // }
  
  var border_range = sheet.getRange(1, 1, last_non_empty_row + 3, 3);
  border_range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1, 1, 2, 1).setFontWeight("bold");
  sheet.getRange(3,1,1,3).setFontWeight("bold");
  sheet.getRange(last_non_empty_row + 5, 1).setValue("Region").setFontWeight("bold");
  var last_blocked_range = sheet.getRange(last_non_empty_row + 4,1,1,26).setBackground("orange");
  sheet.getRange(last_non_empty_row + 5, 1).setValue("Cost Model").setFontWeight("bold");
  
  //var last_row_protection = last_blocked_range.protect().setDescription('Last non Empty Row restriction before cost function area');
  //last_row_protection.addEditor('service-account@spreadsheet-rajabose.iam.gserviceaccount.com');
  //last_row_protection.addEditors(['service-account@spreadsheet-rajabose.iam.gserviceaccount.com']);
  
  //last_row_protection.removeEditors(last_row_protection.getEditors());
  //if (last_row_protection.canDomainEdit()) {
  //  last_row_protection.setDomainEdit(false);
 // }
  
  //var first_columm_range = sheet.getRange(1, 1, last_non_empty_row + 3, 1);
  //var first_column_range_protection = first_columm_range.protect().setDescription('First Column restriction');
  //first_column_range_protection.addEditor('service-account@spreadsheet-rajabose.iam.gserviceaccount.com');
  //first_column_range_protection.addEditors(['service-account@spreadsheet-rajabose.iam.gserviceaccount.com']);
  
  //first_column_range_protection.removeEditors(first_column_range_protection.getEditors());
  //if (first_column_range_protection.canDomainEdit()) {
  //  first_column_range_protection.setDomainEdit(false);
 //}  
}

function protect_sheet(sheet){
  sheet.activate();
  Number.prototype.noExponents= function(){
    var data= String(this).split(/[eE]/);
    if(data.length== 1) return data[0]; 

    var  z= '', sign= this<0? '-':'',
    str= data[0].replace('.', ''),
    mag= Number(data[1])+ 1;

    if(mag<0){
        z= sign + '0.';
        while(mag++) z += '0';
        return z + str.replace(/^\-/,'');
    }
    mag -= str.length;  
    while(mag--) z += '0';
    return str + z;
  }
  
  var last_non_empty_row = sheet.getRange("A4:A").getValues().filter(String).length;
  var sheet_id  = SpreadsheetApp.getActiveSheet().getSheetId();
  var sheet_id = sheet_id.noExponents();
  create_protect_ranges(last_non_empty_row + 2, 0, last_non_empty_row + 3, 26, sheet_id);
  create_protect_ranges(2, 0, last_non_empty_row + 4, 1, sheet_id);
  create_protect_ranges(2, 0, 3, 3, sheet_id);
}

function populate_specification(sheet, response){
  var count = 4;
  var ams_response = response["ams_attribute_fields"];
  var specification_list = [{"primary_text_fields": "populate_textbox_fields"},
                            {"primary_checkbox_fields":"populate_checkbox_fields"}, 
                            {"primary_dropdown_fields": "populate_dropdown_fields"}, 
                            {"secondary_text_fields": "populate_textbox_fields"},
                            {"secondary_checkbox_fields": "populate_checkbox_fields"}, 
                            {"secondary_dropdown_fields": "populate_dropdown_fields"}
                           ];
  
  var ams_specitification_list = [{"dropdown": "populate_ams_dropdown_fields"},
                                      {"checkbox": "populate_ams_checkbox_fields"},
                                      {"text":"populate_ams_textbox_fields"}
                                     ];

  //for(var spec in specification_list){
  //  if(response.hasOwnProperty(Object.keys(specification_list[spec])[0]) && response[Object.keys(specification_list[spec])[0]].length > 0){
  //    eval(specification_list[spec][Object.keys(specification_list[spec])[0]])(sheet, response[Object.keys(specification_list[spec])[0]], count);
  //    count = count + response[Object.keys(specification_list[spec])[0]].length;
  //  }
 // }
  
  ams_response.push({
            "attribute_value": {},
            "attribute": {
                "attribute_input_type": "dropdown",
                "attribute_value_options": [
                    "East",
                    "West",
                    "North",
                    "South"
                ],
                "mandatory": false,
                "is_capability_attribute": false,
                "active": true,
                "priority": 1,
                "specification_type": "primary",
                "data_type": "list_of_strings",
                "help_text": "",
                "name": "Region",
                "master_unit_id": null,
                "master_units": []
            }
        });
  for(var row in ams_response){
    for(var spec in ams_specitification_list){
      if(ams_response[row]["attribute"]["attribute_input_type"] === Object.keys(ams_specitification_list[spec])[0]){
      eval(ams_specitification_list[spec][Object.keys(ams_specitification_list[spec])[0]])(sheet, ams_response[row]["attribute"]["name"], ams_response[row]["attribute"]["attribute_value_options"], count);
      count = count + 1;
     }
    }
  }
}
      
function populate_checkbox_fields(sheet, response, count){
  for (var key in response) {
     if (response.hasOwnProperty(key)) {
        var range = sheet.getRange(count, 1);
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        range.setValue(response[key].name);
        sheet.getRange(count, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(response[key].options.split(','), true).build());
        //range = sheet.getRange(count, 3, 1, 3);
        //range.setValues([["No","Checkbox",response[key].options]]);
        count+=1;
     }
   }
}

function populate_ams_checkbox_fields(sheet, attribute_name, attribute_value, count){
  var range = sheet.getRange(count, 1);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  range.setValue(attribute_name);
  sheet.getRange(count, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(attribute_value, true).build());
}

function populate_ams_textbox_fields(sheet, attribute_name, attribute_value, count){
   var range = sheet.getRange(count, 1);
   range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
   range.setValue(attribute_name);
}

function populate_ams_dropdown_fields(sheet, attribute_name, attribute_value, count){
  var range = sheet.getRange(count, 1);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  range.setValue(attribute_name);
  sheet.getRange(count, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(attribute_value, true).build());
}

function populate_textbox_fields(sheet, response, count){
  for (var key in response) {
     if (response.hasOwnProperty(key)) {
        var range = sheet.getRange(count, 1);
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        range.setValue(response[key].name);
        //range = sheet.getRange(count, 3);
        //range.setValue("No");
        count+=1;
     }
   }
}

function populate_dropdown_fields(sheet, response, count){
  for (var key in response) {
     if (response.hasOwnProperty(key)) {
        var range = sheet.getRange(count, 1);
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        range.setValue(response[key].name);
        sheet.getRange(count, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(response[key].options.split(','), true).build());
        //range = sheet.getRange(count, 3, 1, 3);
        //range.setValues([["No","Dropdown",response[key].options]]);
        count+=1;
     }
   }
}

function GoLastRow(activeSheet) {
  activeSheet.getRange('A1:A').createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build();
  var rg = activeSheet.getFilter().setColumnFilterCriteria(1, criteria).getRange();
  var row = rg.getNextDataCell(SpreadsheetApp.Direction.DOWN);  
  LastRow = row.getRow();
  activeSheet.getFilter().remove();
  Logger.log(activeSheet.getRange(LastRow+1, 1).activate());
}

function getPSANDSS(sub_sub_category_id) {
  var headers = {
    "Authorization" : "Token token=9e20e4f6e0fb29fe433e6a3d26dbf2ef",
    "Secret-Token" : "8d982c7d594576d617b32e1a2ac79a960ed63c40352bf24a2098cdb3049ce10714bf6a6e5e79344b33f80dec440ea8d46a2c36c1284367655712d9bf8ec79133"
  };
  var params = {
    "method":"GET",
    "headers":headers
  };
  var url = "https://bizongo.in/api/admin/sub_sub_categories/"+sub_sub_category_id;

  var response = UrlFetchApp.fetch(url, params);
  
  Logger.clear();
  var result = JSON.parse(response);
  return result;
}

function create_sheet(sheet_name){
  var data = {
    "sheet_name": sheet_name
  };
    
  var headers = {
    "Authorization" : "Token token=9e20e4f6e0fb29fe433e6a3d26dbf2ef",
    "Secret-Token" : "8d982c7d594576d617b32e1a2ac79a960ed63c40352bf24a2098cdb3049ce10714bf6a6e5e79344b33f80dec440ea8d46a2c36c1284367655712d9bf8ec79133"
  };
  
  var params = {
    "method":"POST",
    "headers":headers,
    "payload": data
  }
  
  var url = "https://bizongo.in/api/lead-plus/base-products/create-google-sheet";

  var response = UrlFetchApp.fetch(url, params);
  
  Logger.clear();
  var result = JSON.parse(response);
  return result;
}
  
function create_protect_ranges(start_row_index, start_column_index, end_row_index, end_column_index, sheet_id){
  var data = {
    "range": {
         "sheet_id" : sheet_id,
         "end_column_index" : end_column_index,
         "end_row_index" : end_row_index,
         "start_row_index": start_row_index,
         "start_column_index": start_column_index
        },
    "description":"protected ranges from other users"
  };
  
  var headers = {
    "Authorization" : "Token token=9e20e4f6e0fb29fe433e6a3d26dbf2ef",
    "Secret-Token" : "8d982c7d594576d617b32e1a2ac79a960ed63c40352bf24a2098cdb3049ce10714bf6a6e5e79344b33f80dec440ea8d46a2c36c1284367655712d9bf8ec79133"
  };
  var params = {
    "method": "POST",
    "headers": headers,
    'contentType': 'application/json',
    "payload": JSON.stringify(data)
  };
  var url = "https://bizongo.in/api/lead-plus/base-products/create-protect-ranges";
  var staging_url = "https://qa1.indopus.in/api/lead-plus/base-products/create-protect-ranges";

  var response = UrlFetchApp.fetch(url, params);
  
  var result = JSON.parse(response);
  return result;
}

function get_product_specification_cell(list, specification){
  for(var key in list){
    if (list[key][0] === specification){
      return parseInt(key) + 4;
    }
  }
}

function getCategoryNames(){
  //var spreadsheet = SpreadsheetApp.getActive();
  //var sheet = spreadsheet.getSheetByName('Data for bizongo.com/tools/price-calculator');
  //sheet.activate();
  //var last_non_empty_row = sheet.getRange("A2:A").getValues().filter(String).length;
  //var category_list = sheet.getRange(2, 1, last_non_empty_row, 2).getValues();
  var response = getSSCategory();
  var category_list = response.sub_sub_categories;
  return category_list;
}

function createCategorySheet() {
  //Open a dialog
  var htmlDlg = HtmlService.createTemplateFromFile('category_list').evaluate()
      .setWidth(200)
      .setHeight(150);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'SS Category List');
}

function getSSCategory(){
  var headers = {
    "Authorization" : "Token token=9e20e4f6e0fb29fe433e6a3d26dbf2ef",
    "Secret-Token" : "8d982c7d594576d617b32e1a2ac79a960ed63c40352bf24a2098cdb3049ce10714bf6a6e5e79344b33f80dec440ea8d46a2c36c1284367655712d9bf8ec79133"
  };
  var params = {
    "method":"GET",
    "headers":headers
  };
  var url = "https://bizongo.in/api/admin/sub_sub_categories/all";

  var response = UrlFetchApp.fetch(url, params);
  
  Logger.clear();
  var result = JSON.parse(response);
  return result;
}

function rsc_corrugated_box(){
 return {
    "id": 260,
    "name": "PP Woven Box",
    "image": "Corrugated_box.jpg",
    "visibility": true,
    "meta_keywords": "",
    "description": {
        "body": "This is a corrugated category where price is suggested by price model"
    },
    "priority_order": null,
    "gst_percentage": 12,
    "hsn_number": "48191010",
    "marketing_fees": 10,
    "tags": "Multi-purpose Corrugated Boxes;Corrugated Gift Boxes;Corrugated Shoe Boxes;Corrugated Bottle Holders;Corrugated Milk Boxes;Corrugated Jewellary Boxes;Corrugated Dress Boxes;Corrugated Wine Boxes;Corrugated Chocolate Boxes;Duplex Corrugated Boards;Corrugated Angles;Corrugated Boardes;Flute Boardes;Straw Boards;Flute Boxes;Kraft Papers",
    "return_policy": "NA",
    "created_at": "2019-06-28T00:17:19.138+05:30",
    "updated_at": "2019-08-23T16:35:35.981+05:30",
    "created_by": null,
    "updated_by": "pradeep.m@bizongo.com",
    "sub_category_name": "Boxes And Cartons",
    "sub_category_id": 23,
    "delivery_region": "All over India",
    "ams_attributes_exists": false,
    "primary_text_fields": [
        {
            "name": "Total GSM",
            "optional": true,
            "priority": 15,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Height (In mm)",
            "optional": false,
            "priority": 1,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Length (In mm)",
            "optional": false,
            "priority": 2,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": {
                "is_enabled": {
                    "true": false
                }
            },
            "tolerance": null
        },
        {
            "name": "Breadth (In mm)",
            "optional": false,
            "priority": 3,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": {
                "is_enabled": {
                    "true": false
                }
            },
            "tolerance": null
        },
        {
            "name": "Bursting Factor",
            "optional": false,
            "priority": 7,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Moisture content (%)",
            "optional": false,
            "priority": 17,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Paper Specs (In GSM)",
            "optional": false,
            "priority": 6,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Weight of the CFC (In gm)",
            "optional": false,
            "priority": 8,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": {
                "is_enabled": {
                    "true": false
                }
            },
            "tolerance": null
        },
        {
            "name": "Bursting Strength (In Kg/cm2)",
            "optional": false,
            "priority": 8,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Box Compression Strength (In KGF)",
            "optional": true,
            "priority": 11,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        }
    ],
    "primary_checkbox_fields": [],
    "primary_dropdown_fields": [
        {
            "name": "Printing",
            "options": "1 Color,2 Color,3 Color,4 color,5 Color,Multicolor,No Printing,offset,Flexo",
            "optional": false,
            "priority": 13,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "No of Ply",
            "options": "3,5,7",
            "optional": false,
            "priority": 4,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Packaging",
            "options": "Yes,No",
            "optional": true,
            "priority": 1,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Seperator",
            "options": "Yes,No",
            "optional": false,
            "priority": 14,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Lamination",
            "options": "Yes,No",
            "optional": true,
            "priority": 1,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "No Of Colors",
            "options": "No Color,1 Color,2 Color,3 Color,4 Color,5 Color,6 Color,7 Color,8 Color",
            "optional": true,
            "priority": 1,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Type of Joint",
            "options": "Glued,Stapled,Cloth masking",
            "optional": false,
            "priority": 16,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Type of flute",
            "options": "A,B,C",
            "optional": false,
            "priority": 5,
            "search_show": false,
            "index": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        }
    ],
    "secondary_text_fields": [
        {
            "name": "Dead weight",
            "optional": true,
            "priority": 1,
            "search_show": false,
            "index": false,
            "number": true,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Product Code",
            "optional": true,
            "priority": 19,
            "search_show": false,
            "index": false,
            "number": false,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        },
        {
            "name": "Volumetric weight",
            "optional": true,
            "priority": 1,
            "search_show": false,
            "index": false,
            "number": true,
            "show_invoice": false,
            "coa_specification": null,
            "tolerance": null
        }
    ],
    "secondary_checkbox_fields": [],
    "secondary_dropdown_fields": [],
    "image_url": "https://bizongo-staging-1.s3.amazonaws.com/uploads/picture/image/53251/Integra_Crates_With_Attached_lids.JPG",
    "matrix_attributes": [
        {
            "id": 447,
            "name": "Pack Size",
            "units": "Boxes",
            "use_for_price_calculation": true,
            "unit_for_price_calculation": "Boxes"
        }
    ]
};
}


