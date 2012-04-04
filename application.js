/////////////////////// metadata uploader ///////////////////////////////
/////////////////////////////////////////////////////////////////////////

//MENU

function onOpen() {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is executed.
  menuEntries.push({name: "Upload dataset", functionName: "make_form"});
  menuEntries.push(null);
  menuEntries.push({name: "Get datasets", functionName: "get_datasets"});
  menuEntries.push({name: "Update datasets", functionName: "update_datasets"});
  menuEntries.push({name: "Make dataset template", functionName: "make_dataset_template"});

  ss.addMenu("CKAN", menuEntries);
}

function make_dataset_template() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  if (sheet !== null) {
    Browser.msgBox("Dataset sheet already present");
    return
  }
  sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('dataset');
  sheet.getRange(1,1).setValue('apikey');
  sheet.getRange(1,2).setValue('error');
  sheet.getRange(1,3).setValue('name');
  sheet.setFrozenRows(1);
}
 
//updates dataset in sheet

function update_datasets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  extend_sheet_()  
  for (var i = 2; i < 1000; ++i) {
    if(sheet.getRange(i,1).getValue() === ''){break;};
    row_to_ckan_(i);
  }
}

function get_datasets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  extend_sheet_()  
  for (var i = 2; i < 1000; ++i) {
    if(sheet.getRange(i,1).getValue() === ''){break;};
    ckan_to_row_(i);
  }
}




function ckan_to_row_(row_num) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset')
  var header_row = sheet.getRange("1:1").getValues()[0];
  var row = sheet.getRange(row_num + ":" + row_num).getValues()[0];

  var headers = {'Content Type': 'application/json'};
  headers.authorization = get_apikey_from_row_(header_row, row);
  var url = get_ckanurl_from_row_(header_row, row);
  var id_or_name = get_id_or_name_from_row_(header_row, row);

  var get_json = Utilities.jsonStringify({"id": id_or_name});
  
  try {
    var response = UrlFetchApp.fetch(url + 'api/action/package_show',
                                    {"headers": headers,
                                    "method": "post",
                                    "payload": get_json});
  } catch(err) {
    var ckan_error = err.message.match('{.*$');
    if (ckan_error.length !== 1) {throw err;};
    var error_obj = Utilities.jsonParse(ckan_error[0]);
    if (error_obj.error.__type !== 'Not Found Error'){
      var error_message = Utilities.jsonStringify(error_obj.error);
      error_to_row_(row_num, error_message);
      return;
    };
  };
  
  error_to_row_(row_num, '');
  var result = Utilities.jsonParse(response.getContentText()).result
  var flattened_result = flatten_object_(result);
  update_row_from_dataset_(row_num, flatten_object_(result));
}


function row_to_ckan_(row_num) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset')
  var header_row = sheet.getRange("1:1").getValues()[0];
  var row = sheet.getRange(row_num + ":" + row_num).getValues()[0];
  var headers = {'Content Type': 'application/json'};
  headers.authorization = get_apikey_from_row_(header_row, row);
  var dataset = get_dataset_from_row(row_num);
  var url = get_ckanurl_from_row_(header_row, row);
  var id_or_name = get_id_or_name_from_row_(header_row, row);
  dataset.id = id_or_name
  var json_dataset = Utilities.jsonStringify(dataset);

  var get_json = Utilities.jsonStringify({"id": id_or_name});
  var action_call = 'package_update'
  
  try {
    var response = UrlFetchApp.fetch(url + 'api/action/package_show',
                                    {"headers": headers,
                                    "method": "post",
                                    "payload": get_json});
  } catch(err) {
    var ckan_error = err.message.match('{.*$');
    if (ckan_error.length !== 1) {throw err;};
    var error_obj = Utilities.jsonParse(ckan_error[0]);
    if (error_obj.error.__type !== 'Not Found Error'){
      var error_message = Utilities.jsonStringify(error_obj.error);
      error_to_row_(row_num, error_message);
      return;
    };
    action_call = 'package_create';
    delete dataset.id;
  };  
  
  var json_dataset = Utilities.jsonStringify(dataset);   
  
  try {
    var response = UrlFetchApp.fetch(url + 'api/action/' + action_call,
                                    {"headers": headers,
                                    "method": "post",
                                    "payload": json_dataset});
  } catch(err) {
    var ckan_error = err.message.match('{.*$');
    if (ckan_error === null || ckan_error === undefined) {throw err};
    var error_obj = Utilities.jsonParse(ckan_error[0]);
    var error_message = Utilities.jsonStringify(error_obj.error);
    error_to_row_(row_num, error_message);
    return;
  };
  
  error_to_row_(row_num, '');
  var result = Utilities.jsonParse(response.getContentText()).result
  var flattened_result = flatten_object_(result);
  update_row_from_dataset_(row_num, flatten_object_(result));
}

function error_to_row_(row_num, error){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  var header_row = sheet.getRange("1:1").getValues()[0];
  for (var i = 0; i < header_row.length; ++i) {
    if (header_row[i] === 'error') {
      sheet.getRange(row_num,i+1).setValue(error);
      break;
    };
  }; 
}


function update_row_from_dataset_(row_num, dataset){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  var header_row = sheet.getRange("1:1").getValues()[0];
  
  var header_set = {};
  for (var i = 0; i < header_row.length; ++i) {
    header_set[header_row[i]] = '';
  };
  
  for (var key in dataset) {
    Logger.log(key)
    if (header_set[key] === undefined) {
      add_header_row_(key);
    };
  };
  
  var header_row = sheet.getRange("1:1").getValues()[0];  
  var row = sheet.getRange(row_num + ":" + row_num).getValues()[0]; 
  
  for (var i = 0; i < header_row.length; ++i) {
    var cur_val = row[i];
    if (cur_val !== '') {continue;};
    var header = header_row[i];
    if (dataset[header] !== undefined){
      sheet.getRange(row_num,i+1).setValue(dataset[header]);
    };
  }; 
}

function add_header_row_(key){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset')
  var header_row = sheet.getRange("1:1").getValues()[0];
  for (var i = 0; i < header_row.length; ++i) {
    if (header_row[i] === '') {
      sheet.getRange(1,i+1).setValue(key);
      break;
    };
  }
}
      
  

function get_dataset_from_row(row_num) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  var header_row = sheet.getRange("1:1").getValues()[0];
  var row = sheet.getRange(row_num + ":" + row_num).getValues()[0];
  var dataset = {};

  for (var i = 0; i < header_row.length; ++i) {
    var header = header_row[i];
    if (header === '') {break;};
    var cell_data = row[i];
    if (header === 'apikey' || header === 'error' ||  header === 'ckanurl') {
      continue;
    };
    if (cell_data === '') {continue;};
    
    var header_items_ = header.split('__');
    if (header_items_.length === 1){
      dataset[header] = row[i];
      continue;
    };
    
    var section = header_items_[0];
    var order = header_items_[1];
    var key = header_items_[2];
     
    
    if (dataset[section] === undefined){
        dataset[section] = [];
    }
    if (dataset[section][order] === undefined){
      dataset[section][order] = {};
    };
    dataset[section][parseInt(order)][key] = row[i];
  }
 
  return dataset
}


function get_apikey_from_row_(header_row, row) {
  for (var i = 0; i < header_row.length; ++i) {
    var header = header_row[i];    
    if (header === 'apikey') {
      apikey = row[i];
      if (apikey === '') {
        throw 'Api key not found for row ' + (i+2);
      }
      return apikey
    };
  };
  throw 'No apikey header found'
}


function get_ckanurl_from_row_(header_row, row) {
  
  for (var i = 0; i < header_row.length; ++i) {
    var header = header_row[i];    
    if (header === 'ckanurl') {
      url = row[i];
      if (url === '') {
        throw 'ckanurl not found for row ' + (i+2);
      }
      return url
    };
  };
  return 'http://www.thedatahub.org/'
}  


function get_id_or_name_from_row_(header_row, row) {
  for (var i = 0; i < header_row.length; ++i) {
    var header = header_row[i];    
    if (header === 'id') {
      id_or_name = row[i];
      if (id_or_name !== '') {
          return id_or_name
      }
    };
  };  
  for (var i = 0; i < header_row.length; ++i) {
    var header = header_row[i];    
    if (header === 'name') {
      id_or_name = row[i];
      if (id_or_name !== '') {
          return id_or_name
      }
    };
  };
  throw 'dataset id or name not found for row ' + (i+2);
}

/////////////////////////  WEBSTORE UPLOAD //////////////////////////////////
/////////////////////////////////////////////////////////////////////////////


///FORM

function make_form(values, errors) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var app = UiApp.createApplication().setTitle('Upload data form');
  
  var empty_form_data = {"ckan": "", "webstore": "",
                         "apikey" : "", "user_name": "", 
                         "dataset_name": "", "resource_name": ""};
    
  if (values === undefined || values === {}) {
    var values = _.extend({}, empty_form_data);
    values.ckan = 'http://thedatahub.org/'
    values.dataset_name = ss.getName();
    values.resource_name = SpreadsheetApp.getActiveSpreadsheet().getSheetName()
  }

    
  if (errors === undefined) {
    var errors = {};
  }
  errors = _.extend(empty_form_data, errors)
    
  Logger.log(errors)
  
  var grid = app.createGrid(6, 3);
  grid.setWidget(0, 0, app.createLabel('CKAN instance:'));
  grid.setWidget(0, 1, app.createTextBox().setText(values.ckan).setName('ckan'));
  grid.setWidget(0, 2, app.createLabel(errors.ckan).setStyleAttribute('color', 'red'));
  
  grid.setWidget(1, 0, app.createLabel('Apikey:'));
  grid.setWidget(1, 1, app.createTextBox().setText(values.apikey).setName('apikey'));
  grid.setWidget(1, 2, app.createLabel(errors.apikey).setStyleAttribute('color', 'red'));
  
  grid.setWidget(2, 0, app.createLabel('Dataset name:'));
  grid.setWidget(2, 1, app.createTextBox().setText(values.dataset_name).setName('dataset_name'));
  grid.setWidget(2, 2, app.createLabel(errors.dataset_name).setStyleAttribute('color', 'red'));
  
  grid.setWidget(3, 0, app.createLabel('Resource name:'));
  grid.setWidget(3, 1, app.createTextBox().setText(values.resource_name).setName('resource_name'));
  grid.setWidget(3, 2, app.createLabel(errors.resource_name).setStyleAttribute('color', 'red'));

  
  // Create a vertical panel..
  var panel = app.createVerticalPanel();
  
  // ...and add the grid to the panel
  panel.add(grid);
  
  // Create a button and click handler; pass in the grid object as a callback element and the handler as a click handler
  // Identify the function b as the server click handler

  var button = app.createButton('submit');
  var handler = app.createServerClickHandler('upload_data_');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);
  
  // Add the button to the panel and the panel to the application, then display the application app in the Spreadsheet doc
  panel.add(button);
  app.add(panel);
  ss.show(app);
}


/// FORM HANDLER

function upload_data_(form_data) {
  
  Logger.log('parameter')  
  Logger.log(form_data.parameter)
  
  var form_data = {'ckan' : form_data.parameter.ckan,
                   'apikey': form_data.parameter.apikey,
                   'dataset_name': form_data.parameter.dataset_name,
                   'resource_name': form_data.parameter.resource_name};
  
  Logger.log('form_data')  
  Logger.log(form_data)
  
  ckan_url = form_data.ckan;
  
  var errors = {};
  
  if (form_data.apikey === '') {errors.apikey = 'No Api key supplied'};
  if (form_data.ckan === '') {errors.ckan = 'No ckan url supplied'};
  if (form_data.dataset_name === '') {errors.dataset_name = 'No dataset name supplied'};
  if (form_data.resource_name === '') {errors.resource_name = 'No resource name supplied'};
  
  Logger.log('errors');
  Logger.log(errors);
  
  if (!_.isEmpty(errors)) {
      var app = UiApp.getActiveApplication();
      app.close();
      make_form(form_data, errors);
      return
  };
  
      
  create_package_if_new_(form_data);
  dataset = get_from_ckan_(form_data);
  
  // find resource_id of named resource
  
  var resource_id;
  for(var i=0; i<dataset.resources.length; i++) {
    var resource_name = dataset.resources[i].name;
    if (resource_name === form_data.resource_name) {
      dataset.resources[i].webstore_url = 'enabled';
      resource_id = dataset.resources[i].id
      break;
    }
  };
  
  // create new resource if not found
  if(!resource_id) {
    // create the resource
    dataset.resources.push({'name': form_data.resource_name,
                            'url': 'internal',
                            'webstore_url': 'enabled'});
  }
  
  // post changes back
  dataset = post_to_ckan_(dataset, form_data);
  
  // refech resourse_id, should be there now.
  for(var i=0; i<dataset.resources.length; i++) {
    var resource_name = dataset.resources[i].name;
    if (resource_name === form_data.resource_name) {
      resource_id = dataset.resources[i].id
      break;
    }
  };  
    
  // Upload the currently active sheet to the webstore
  var sheetname = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  var auth = {'apikey': form_data.apikey, 'ckan': form_data.ckan};
  var resource = {'id': resource_id};
  Logger.log('uploading');
  upload_sheet_to_webstore(sheetname, auth, resource);
  Logger.log('finished uploading');
  
  // Close the form
  var app = UiApp.getActiveApplication();
  app.close();
  // The following line is REQUIRED for the widget to actually close.
  return app;
}



function create_package_if_new_(form_data) {
  Logger.log(form_data)
  var json_dataset = Utilities.jsonStringify({"name": munge_name(form_data.dataset_name),
                                              "title": form_data.dataset_name});
  var headers = {'Authorization': form_data.apikey,
                 'Accept': 'application/json'};
  
  try {
    var response = UrlFetchApp.fetch(ckan_root_uri_() + 'api/action/package_create',
                                    {"headers": headers,
                                     "contentType": "application/json",
                                     "method": "post",
                                     "payload": json_dataset});
    return Utilities.jsonParse(response.getContentText()).result
  } catch(err) {
    var ckan_error = err.message.match('{.*$');
    Logger.log('ckan_error');
    Logger.log(ckan_error);
    if (ckan_error === null || ckan_error === undefined) {throw err};
    var error_obj = Utilities.jsonParse(ckan_error[0]);
    if (error_obj.success == false) {
      if (error_obj.error.name !== undefined && error_obj.error.name[0] !== "That URL is already in use.") {
         throw err;
      }
    return false
    }

  };
};​

/**
 * Download the dataset, identified by name, from CKAN.
 * 
*  - result
 */
function get_from_ckan_(result) {
  var headers = {'Authorization': result.apikey,
                 'Accept': 'application/json'};
  var get_json = Utilities.jsonStringify({"id": munge_name(result.dataset_name)});
  
  var response = UrlFetchApp.fetch(ckan_root_uri_() + 'api/action/package_show',
                                    {"headers": headers,
                                    "method": "post",
                                    "payload": get_json});
  
  return Utilities.jsonParse(response.getContentText()).result;
}
  
function test_create_package() {

  var result = {'apikey': '61b5d189-1125-43e5-8759-92af202b0820',
                'user_name': 'raz',
                'dataset_name': 'test_google_docs7',
                'resource_name': 'test_google_docs7_resource'};
  create_package_if_new_(result);
}
  
function ckan_root_uri_() {
  return ckan_url;
}

/**
 * Update the given dataset
 */
function post_to_ckan_(dataset, form_data) {
  var headers = {'Authorization': form_data.apikey,
                 'Accept': 'application/json'};
  var url = ckan_root_uri_();
  var json_dataset = Utilities.jsonStringify(dataset);
  
  var response = UrlFetchApp.fetch(url + 'api/action/package_update',
                                  {"headers": headers,
                                   "contentType": "application/json",
                                   "method": "post",
                                   "payload": json_dataset});
  
  return Utilities.jsonParse(response.getContentText()).result
}



/*
 * Uploads the data in the given sheet to the webstore.
 * 
 * - sheetname is the name of the sheet that contains the data
 *   to upload to the webstore.
 * - auth is an oject containing the authorization info, namely:
 *    apikey : the api key
 *    name   : the user name
 * - resource is an object containing information about where
 *   to upload the data to, namely:
 *    database : the database name
 *    table : the table name
 */
function upload_sheet_to_webstore(sheetname, auth, resource) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var column_names = get_column_names_(sheet);
  var data = get_data_(sheet, column_names);
  
  //try and delete the data first
  var response = UrlFetchApp.fetch(auth.ckan + 'api/data/' + resource.id,
                                  {"contentType": "application/json",
                                   "headers": headers,
                                   "method": "delete"}); 
  
    
  var headers = {
    'authorization': auth.apikey,
    'Accept': "application/json"};
  var url = auth.ckan + 'api/data/' + resource.id + '/_bulk';
  
  var bulk_data = []
  
  for (var i = 0; i < data.length; ++i){
    bulk_data.push(Utilities.jsonStringify({"index": {"_id": i+1}}));
    bulk_data.push(Utilities.jsonStringify(data[i]));
  }
  
  var payload = bulk_data.join('\n') + '\n';
  
  Logger.log('url');
  Logger.log(url);

  var response = UrlFetchApp.fetch(url,
                                   {"contentType": "application/json",
                                    "headers": headers,
                                    "method": "post",
                                    "payload": payload});
  var s = response.getContentText();
  Logger.log(s);
}

function delete_table_on_webstore_(tablename, auth, resource) {
  var headers = {
    'Accept': 'application/json',
    'authorization': auth.apikey};
  var url = resource_url_(resource);
  
  try {
    var response = UrlFetchApp.fetch(url,
                                     {"contentType": "application/json",
                                      "headers": headers,
                                      "method": "delete"});
  } catch(err) { // On success, the webstore returns a 410 which is caught here.
    
    debugger;
    var error_message = err.message;
    var s = "";
  };
}







/////////////////////////////////  UTILS ///////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

// make sure there are enough columns

function extend_sheet_(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dataset');
  var length = sheet.getRange("1:1").getValues()[0].length;
  if (sheet.getMaxColumns() < 200){
    sheet.insertColumnsAfter(length, 200 - sheet.getMaxColumns());
  }
}


function munge_name(slug, delimiter) {
  Logger.log('slug')
  Logger.log(slug)

  if (delimiter == undefined) {
    var delimiter = '-';
  }
  
  var regexToHyphen = [new RegExp('[ .:/_]', 'g'), 
                       new RegExp('[^a-zA-Z0-9-_]', 'g'), 
                       new RegExp('-+', 'g')];
  
  var regexToDelete = [new RegExp('^-*', 'g'), 
                       new RegExp('-*$', 'g')];
  
  _.each(regexToHyphen, function(regex) { slug = slug.replace(regex, delimiter); });
  _.each(regexToDelete, function(regex) { slug = slug.replace(regex, ''); });
  
  slug=slug.substring(0,200);

  return slug.toLowerCase();
}


    
function keys_(obj){
  var keys = [];
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      keys.push(key);
    }
  }
  return keys;
}
function items_(obj){
  var items = [];
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      items.push({"key": key, "value": obj[key]});
    }
  }
  return items;
}



function flatten_object_(obj){
  var new_obj = {};
  var obj_items = items_(obj);
  for (var i = 0; i < obj_items.length; ++i){
    var key = obj_items[i].key;
    var value = obj_items[i].value;
    if (value === null){continue;};
    if (value.constructor != Array){
      new_obj[key] = value;
      continue;
    }
    for (var j = 0; j < value.length; ++j){
      var sub_obj_items = items_(value[j]);
      for (var k = 0; k < sub_obj_items.length; ++k){
        if (sub_obj_items[k].value=== null){continue;};
        new_obj[key + '__' + j + '__' + sub_obj_items[k].key] = sub_obj_items[k].value;
      }
    }
  }
  return new_obj;
}


/*
 * Returns an ordered list of column names found in the given sheet.
 * 
 *  - The first row defines the column names.
 *  - Drops any trailing columns that have blank names
 */
function get_column_names_(sheet) {
  var header_row = sheet.getRange("1:1").getValues()[0];
  var last_non_blank_index = last_index_of_(header_row, function(element){return element != "";});
  if (last_non_blank_index == -1) { return []; }
  return header_row.slice(0, last_non_blank_index+1);
}

/**
 * Returns the index of the last element that satisfies f
 * If no element satisfies f, -1 is returned.
 */
function last_index_of_(ary, f) {
  for (var i = ary.length-1; i >= 0; i--) {
    if ( f(ary[i]) ) { return i; }
  };
  return -1;
}

/**
 * Pulls out the data from the spreadsheet, and returns it as
 * a list of dicts.
 *
 *  - each element of the returned list is a dict mapping column
 *    names to column values
 *  - ignores any empty column names
 */
function get_data_(sheet, column_names) {
  var data = [];
  var max_row = sheet.getLastRow();
  for (var row_num = 2; row_num <= max_row; row_num++) {  // indexed from 1, skip header row
    var row = sheet.getRange(row_num + ":" + row_num).getValues();
    var row_data = {};
    for (var column_index = 0; column_index < column_names.length; column_index++) {
      var column_name = column_names[column_index];
      if ( column_name != "" ){
        row_data[column_name] = row[0][column_index];
      }
    };
    data.push(row_data);
  };
  return data;
}
