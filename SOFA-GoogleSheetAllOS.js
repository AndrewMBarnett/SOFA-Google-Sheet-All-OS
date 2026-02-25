function onOpen() {
  Logger.log('Creating custom menu.');
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OS Updates Menu')
    .addItem('Fetch macOS Data', 'fetchJsonAndPopulateSheet')
    .addItem('Fetch iOS Data', 'fetchiOSJsonAndPopulateSheet')
    .addItem('Fetch tvOS Data', 'fetchtvOSJsonAndPopulateSheet')
    .addItem('Fetch watchOS Data', 'fetchwatchOSJsonAndPopulateSheet')
    .addItem('Fetch visionOS Data', 'fetchvisionOSJsonAndPopulateSheet')
    .addItem('Fetch Safari Data', 'fetchSafariJsonAndPopulateSheet')
    .addSeparator()
    .addItem('Fetch All Platforms', 'fetchAllPlatforms')
    .addItem('Update Dashboard Only', 'updateDashboard')
    .addToUi();
}

function fetchAllPlatforms() {
  Logger.log('Fetching all platform data.');
  fetchJsonAndPopulateSheet();
  fetchiOSJsonAndPopulateSheet();
  fetchtvOSJsonAndPopulateSheet();
  fetchwatchOSJsonAndPopulateSheet();
  fetchvisionOSJsonAndPopulateSheet();
  fetchSafariJsonAndPopulateSheet();
  updateDashboard();
  SpreadsheetApp.getUi().alert('All platforms and dashboard updated successfully!');
}

function fetchBothSheets() {
  fetchAllPlatforms();
}

function updateDashboard() {
  Logger.log('Starting updateDashboard function.');
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    var dashboard = spreadsheet.getSheetByName('Dashboard');
    if (!dashboard) {
      dashboard = spreadsheet.insertSheet('Dashboard', 0);
    }
    
    dashboard.clear();
    SpreadsheetApp.flush();
    
    var timestamp = new Date();
    var timestampText = 'Last Updated: ' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss z");
    dashboard.getRange(1, 1).setValue(timestampText);
    dashboard.getRange(1, 1).setFontWeight('bold').setFontSize(11).setBackground('#FFF3CD');
    
    var headers = ["Platform", "OS Version", "Latest Update", "Product Version", "Release Date", "CVE Count", "Actively Exploited CVEs", "Days Since Previous", "Security Info", "SOFA Details"];
    dashboard.getRange(2, 1, 1, headers.length).setValues([headers]);
    dashboard.setFrozenRows(2);
    
    var headerRange = dashboard.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    SpreadsheetApp.flush();
    
    var sofaLinks = getSofaLinksMapping();
    
    var dashboardData = [];
    
    var platforms = [
      {url: 'https://sofa.macadmins.io/v2/macos_data_feed.json', name: 'macOS', color: '#E8F0FE'},
      {url: 'https://sofafeed.macadmins.io/v2/ios_data_feed.json', name: 'iOS', color: '#FFF3E0'},
      {url: 'https://sofa.macadmins.io/v2/tvos_data_feed.json', name: 'tvOS', color: '#F3E5F5'},
      {url: 'https://sofa.macadmins.io/v2/watchos_data_feed.json', name: 'watchOS', color: '#E0F2F1'},
      {url: 'https://sofa.macadmins.io/v2/visionos_data_feed.json', name: 'visionOS', color: '#FCE4EC'},
      {url: 'https://sofa.macadmins.io/v2/safari_data_feed.json', name: 'Safari', color: '#E3F2FD'}
    ];
    
    platforms.forEach(function(platform) {
      try {
        var platformData = fetchLatestVersions(platform.url, platform.name, sofaLinks);
        dashboardData = dashboardData.concat(platformData);
      } catch (e) {
        Logger.log('Error fetching ' + platform.name + ' dashboard data: ' + e);
      }
    });
    
    if (dashboardData.length > 0) {
      dashboard.getRange(3, 1, dashboardData.length, headers.length).setValues(dashboardData);
      SpreadsheetApp.flush();
      
      var dataRange = dashboard.getRange(3, 1, dashboardData.length, headers.length);
      dataRange.setHorizontalAlignment('left');
      
      platforms.forEach(function(platform) {
        for (var i = 0; i < dashboardData.length; i++) {
          if (dashboardData[i][0] === platform.name) {
            var rowIndex = dashboardData.slice(0, i).filter(function(row) { 
              return row[0] === platform.name; 
            }).length;
            var bgColor = rowIndex % 2 === 0 ? platform.color : '#FFFFFF';
            dashboard.getRange(i + 3, 1, 1, headers.length).setBackground(bgColor); // i + 3 instead of i + 2
          }
        }
      });
      
      SpreadsheetApp.flush();
      
      for (var i = 0; i < dashboardData.length; i++) {
        if (dashboardData[i][6] !== "None" && dashboardData[i][6] !== "") {
          dashboard.getRange(i + 3, 7).setBackground('#F4CCCC').setFontWeight('bold'); // i + 3 instead of i + 2
        }
      }
      
      SpreadsheetApp.flush();
    }
    
    for (var i = 1; i <= headers.length; i++) {
      dashboard.autoResizeColumn(i);
    }
    

    dashboard.setColumnWidth(9, 300);
    dashboard.setColumnWidth(10, 300);
    
    SpreadsheetApp.flush();
    
    Logger.log('Dashboard updated successfully.');
  } catch (e) {
    Logger.log('Error in updateDashboard: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error updating dashboard: ' + e.toString());
    throw e;
  }
}

function addTimestampToSheet(sheet) {
  var timestamp = new Date();
  var timestampText = 'Last Updated: ' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss z");
  sheet.getRange(1, 1).setValue(timestampText);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(11).setBackground('#FFF3CD');
}

function fetchLatestVersions(url, platform, sofaLinks) {
  Logger.log('Fetching latest versions for ' + platform);
  
  try {
    var response = UrlFetchApp.fetch(url);
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching data for ' + platform + ': ' + e);
    return [];
  }
  
  var latestVersions = [];
  
  if (platform === 'Safari' && data.AppVersions && data.AppVersions.length > 0) {
    data.AppVersions.forEach(function(appVersion) {
      if (appVersion.SecurityReleases && appVersion.SecurityReleases.length > 0) {
        var latestRelease = appVersion.SecurityReleases[0];
        
        var activelyExploitedCVEs = latestRelease.ActivelyExploitedCVEs ? latestRelease.ActivelyExploitedCVEs.join(", ") : "None";
        var cveCount = latestRelease.CVEs ? Object.keys(latestRelease.CVEs).length : 0;
        
        var versionMatch = appVersion.AppVersion.match(/Safari (\d+)/);
        var sofaLink = "";
        if (versionMatch && sofaLinks) {
          var key = 'safari_' + versionMatch[1];
          sofaLink = sofaLinks[key] || "";
        }
        
        var row = [
          platform,
          appVersion.AppVersion || "",
          latestRelease.UpdateName || "",
          latestRelease.ProductVersion || "",
          formatDate(latestRelease.ReleaseDate, false),
          cveCount,
          activelyExploitedCVEs,
          latestRelease.DaysSincePreviousRelease != null ? latestRelease.DaysSincePreviousRelease.toString() : "",
          latestRelease.SecurityInfo || "",
          sofaLink
        ];
        
        latestVersions.push(row);
      }
    });
  } else if (data.OSVersions && data.OSVersions.length > 0) {
    data.OSVersions.forEach(function(osVersion) {
      if (osVersion.SecurityReleases && osVersion.SecurityReleases.length > 0) {
        var latestRelease = osVersion.SecurityReleases[0];
        
        var activelyExploitedCVEs = latestRelease.ActivelyExploitedCVEs ? latestRelease.ActivelyExploitedCVEs.join(", ") : "None";
        var cveCount = latestRelease.CVEs ? Object.keys(latestRelease.CVEs).length : 0;
        
        var sofaLink = getSofaLinkForVersion(platform, osVersion.OSVersion, sofaLinks);
        
        var row = [
          platform,
          osVersion.OSVersion || "",
          latestRelease.UpdateName || "",
          latestRelease.ProductVersion || "",
          formatDate(latestRelease.ReleaseDate, false),
          cveCount,
          activelyExploitedCVEs,
          latestRelease.DaysSincePreviousRelease != null ? latestRelease.DaysSincePreviousRelease.toString() : "",
          latestRelease.SecurityInfo || "",
          sofaLink
        ];
        
        latestVersions.push(row);
      }
    });
  }
  
  return latestVersions;
}

function getSofaLinksMapping() {

  var mapping = {
    'macos_26': 'https://sofa.macadmins.io/macos/tahoe',
    'macos_tahoe': 'https://sofa.macadmins.io/macos/tahoe',
    'macos_15': 'https://sofa.macadmins.io/macos/sequoia',
    'macos_sequoia': 'https://sofa.macadmins.io/macos/sequoia',
    'macos_14': 'https://sofa.macadmins.io/macos/sonoma',
    'macos_sonoma': 'https://sofa.macadmins.io/macos/sonoma',
    'macos_13': 'https://sofa.macadmins.io/macos/ventura',
    'macos_ventura': 'https://sofa.macadmins.io/macos/ventura',
    'macos_12': 'https://sofa.macadmins.io/macos/monterey',
    'macos_monterey': 'https://sofa.macadmins.io/macos/monterey',
    
    'ios_26': 'https://sofa.macadmins.io/ios/ios26',
    'ios_18': 'https://sofa.macadmins.io/ios/ios18',
    'ios_17': 'https://sofa.macadmins.io/ios/ios17',

    'safari_26': 'https://sofa.macadmins.io/safari/safari26',
    'safari_18': 'https://sofa.macadmins.io/safari/safari18',
    'safari_17': 'https://sofa.macadmins.io/safari/safari17',
    
    'tvos_26': 'https://sofa.macadmins.io/tvos/tvos26',
    'tvos_18': 'https://sofa.macadmins.io/tvos/tvos18',

    'visionos_26': 'https://sofa.macadmins.io/visionos/visionos26',
    'visionos_2': 'https://sofa.macadmins.io/visionos/visionos2',
    
    'watchos_26': 'https://sofa.macadmins.io/watchos/watchos26',
    'watchos_11': 'https://sofa.macadmins.io/watchos/watchos11'
  };
  
  Logger.log('Using hardcoded SOFA links mapping with ' + Object.keys(mapping).length + ' entries');
  return mapping;
}

function getSofaLinkForVersion(platform, osVersion, sofaLinks) {
  if (!sofaLinks) return "";
  
  var platformKey = platform.toLowerCase();
  
  var keys = [];
  
  var versionMatch = osVersion.match(/(\d+)/);
  if (versionMatch) {
    var majorVersion = versionMatch[1];
    keys.push(platformKey + '_' + majorVersion);
  }
  
  var nameMatch = osVersion.match(/^([A-Za-z]+)/);
  if (nameMatch) {
    var osName = nameMatch[1].toLowerCase();
    keys.push(platformKey + '_' + osName);
  }
  
  var simpleMatch = osVersion.match(/^[A-Za-z]+\s*(\d+)/);
  if (simpleMatch) {
    keys.push(platformKey + '_' + simpleMatch[1]);
  }
  
  for (var i = 0; i < keys.length; i++) {
    var link = sofaLinks[keys[i]];
    if (link) {
      Logger.log('Found SOFA link for ' + platform + ' "' + osVersion + '" using key "' + keys[i] + '": ' + link);
      return link;
    }
  }
  
  Logger.log('No SOFA link found for ' + platform + ' "' + osVersion + '". Tried keys: ' + keys.join(', '));
  return "";
}

function fetchJsonAndPopulateSheet() {
  Logger.log('Starting fetchJsonAndPopulateSheet function.');
  var url = 'https://sofa.macadmins.io/v2/macos_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('macOS');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('macOS');
  }
  
  sheet.clear();
  
  addTimestampToSheet(sheet);
  
  var headers = ["OS Version", "Update Name", "Product Version", "Release Date", "Security Info URL", "CVE Count", "Actively Exploited CVEs", "Days Since Previous Release"];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]); // Row 2 instead of row 1
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Failed to fetch or parse data: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error fetching macOS data: ' + e.toString());
    return;
  }
  
  data.OSVersions.forEach(function(osVersion) {
    Logger.log('Processing OSVersion: ' + osVersion.OSVersion);
    osVersion.SecurityReleases.forEach(function(release) {
      var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
      var row = [
        osVersion.OSVersion || "",
        release.UpdateName || "",
        release.ProductVersion || "",
        formatDate(release.ReleaseDate, false),
        release.SecurityInfo || "",
        release.CVEs ? Object.keys(release.CVEs).length : 0,
        activelyExploitedCVEs,
        release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
      ];
      appendDataToSheet(sheet, [row]);
    });
  });
  
  Logger.log('macOS sheet updated successfully.');
}

function fetchiOSJsonAndPopulateSheet() {
  Logger.log('Starting fetchiOSJsonAndPopulateSheet function.');
  var url = 'https://sofafeed.macadmins.io/v2/ios_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('iOS');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('iOS');
  }
  
  sheet.clear();
  
  addTimestampToSheet(sheet);
  
  var headers = [
    "OS Version",
    "Update Name",
    "Product Version",
    "Release Date",
    "Security Info URL",
    "CVE Count",
    "Actively Exploited CVEs",
    "Days Since Previous Release"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching/parsing data: ' + e);
    SpreadsheetApp.getUi().alert('Error fetching iOS data: ' + e.toString());
    return;
  }
  
  data.OSVersions.forEach(function(osVersion) {
    Logger.log('Processing OSVersion: ' + osVersion.OSVersion);
    osVersion.SecurityReleases.forEach(function(release) {
      var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
      var row = [
        osVersion.OSVersion || "",
        release.UpdateName || "",
        release.ProductVersion || "",
        formatDate(release.ReleaseDate, false),
        release.SecurityInfo || "",
        release.CVEs ? Object.keys(release.CVEs).length : 0,
        activelyExploitedCVEs,
        release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
      ];
      appendDataToSheet(sheet, [row]);
    });
  });
  
  Logger.log('iOS sheet updated successfully.');
}

function fetchtvOSJsonAndPopulateSheet() {
  Logger.log('Starting fetchtvOSJsonAndPopulateSheet function.');
  var url = 'https://sofa.macadmins.io/v2/tvos_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('tvOS');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('tvOS');
  }
  
  sheet.clear();
  addTimestampToSheet(sheet);
  
  var headers = [
    "OS Version",
    "Update Name",
    "Product Version",
    "Release Date",
    "Security Info URL",
    "CVE Count",
    "Actively Exploited CVEs",
    "Days Since Previous Release"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching/parsing data: ' + e);
    SpreadsheetApp.getUi().alert('Error fetching tvOS data: ' + e.toString());
    return;
  }
  
  data.OSVersions.forEach(function(osVersion) {
    Logger.log('Processing OSVersion: ' + osVersion.OSVersion);
    osVersion.SecurityReleases.forEach(function(release) {
      var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
      var row = [
        osVersion.OSVersion || "",
        release.UpdateName || "",
        release.ProductVersion || "",
        formatDate(release.ReleaseDate, false),
        release.SecurityInfo || "",
        release.CVEs ? Object.keys(release.CVEs).length : 0,
        activelyExploitedCVEs,
        release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
      ];
      appendDataToSheet(sheet, [row]);
    });
  });
  
  Logger.log('tvOS sheet updated successfully.');
}

function fetchwatchOSJsonAndPopulateSheet() {
  Logger.log('Starting fetchwatchOSJsonAndPopulateSheet function.');
  var url = 'https://sofa.macadmins.io/v2/watchos_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('watchOS');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('watchOS');
  }
  
  sheet.clear();
  addTimestampToSheet(sheet);
  
  var headers = [
    "OS Version",
    "Update Name",
    "Product Version",
    "Release Date",
    "Security Info URL",
    "CVE Count",
    "Actively Exploited CVEs",
    "Days Since Previous Release"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching/parsing data: ' + e);
    SpreadsheetApp.getUi().alert('Error fetching watchOS data: ' + e.toString());
    return;
  }
  
  data.OSVersions.forEach(function(osVersion) {
    Logger.log('Processing OSVersion: ' + osVersion.OSVersion);
    osVersion.SecurityReleases.forEach(function(release) {
      var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
      var row = [
        osVersion.OSVersion || "",
        release.UpdateName || "",
        release.ProductVersion || "",
        formatDate(release.ReleaseDate, false),
        release.SecurityInfo || "",
        release.CVEs ? Object.keys(release.CVEs).length : 0,
        activelyExploitedCVEs,
        release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
      ];
      appendDataToSheet(sheet, [row]);
    });
  });
  
  Logger.log('watchOS sheet updated successfully.');
}

function fetchvisionOSJsonAndPopulateSheet() {
  Logger.log('Starting fetchvisionOSJsonAndPopulateSheet function.');
  var url = 'https://sofa.macadmins.io/v2/visionos_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('visionOS');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('visionOS');
  }
  
  sheet.clear();
  addTimestampToSheet(sheet);
  
  var headers = [
    "OS Version",
    "Update Name",
    "Product Version",
    "Release Date",
    "Security Info URL",
    "CVE Count",
    "Actively Exploited CVEs",
    "Days Since Previous Release"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching/parsing data: ' + e);
    SpreadsheetApp.getUi().alert('Error fetching visionOS data: ' + e.toString());
    return;
  }
  
  data.OSVersions.forEach(function(osVersion) {
    Logger.log('Processing OSVersion: ' + osVersion.OSVersion);
    osVersion.SecurityReleases.forEach(function(release) {
      var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
      var row = [
        osVersion.OSVersion || "",
        release.UpdateName || "",
        release.ProductVersion || "",
        formatDate(release.ReleaseDate, false),
        release.SecurityInfo || "",
        release.CVEs ? Object.keys(release.CVEs).length : 0,
        activelyExploitedCVEs,
        release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
      ];
      appendDataToSheet(sheet, [row]);
    });
  });
  
  Logger.log('visionOS sheet updated successfully.');
}

function fetchSafariJsonAndPopulateSheet() {
  Logger.log('Starting fetchSafariJsonAndPopulateSheet function.');
  var url = 'https://sofa.macadmins.io/v2/safari_data_feed.json';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('Safari');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Safari');
  }
  
  sheet.clear();
  addTimestampToSheet(sheet);
  
  var headers = [
    "App Version",
    "Update Name",
    "Product Version",
    "Release Date",
    "Security Info URL",
    "CVE Count",
    "Actively Exploited CVEs",
    "Days Since Previous Release"
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(2);
  
  try {
    var response = UrlFetchApp.fetch(url);
    Logger.log('Fetched data successfully.');
    var json = response.getContentText();
    var data = JSON.parse(json);
  } catch (e) {
    Logger.log('Error fetching/parsing data: ' + e);
    SpreadsheetApp.getUi().alert('Error fetching Safari data: ' + e.toString());
    return;
  }
  
  if (data.AppVersions && data.AppVersions.length > 0) {
    data.AppVersions.forEach(function(appVersion) {
      Logger.log('Processing AppVersion: ' + appVersion.AppVersion);
      if (appVersion.SecurityReleases && appVersion.SecurityReleases.length > 0) {
        appVersion.SecurityReleases.forEach(function(release) {
          var activelyExploitedCVEs = release.ActivelyExploitedCVEs ? release.ActivelyExploitedCVEs.join(", ") : "None";
          var row = [
            appVersion.AppVersion || "",
            release.UpdateName || "",
            release.ProductVersion || "",
            formatDate(release.ReleaseDate, false),
            release.SecurityInfo || "",
            release.CVEs ? Object.keys(release.CVEs).length : 0,
            activelyExploitedCVEs,
            release.DaysSincePreviousRelease != null ? release.DaysSincePreviousRelease.toString() : ""
          ];
          appendDataToSheet(sheet, [row]);
        });
      }
    });
  }
  
  Logger.log('Safari sheet updated successfully.');
}

function appendDataToSheet(sheet, rowData) {
  Logger.log('Appending data to sheet.');
  rowData.forEach(function(row) {
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(lastRow + 1, 1, 1, row.length);
    range.setValues([row]);
    range.setHorizontalAlignment("left");
  });
  Logger.log('Data appended successfully.');
}

function formatDate(dateStr, includeTime) {
  if (!dateStr) return "";
  Logger.log('Formatting date: ' + dateStr);
  try {
    var date = new Date(dateStr);
    if (includeTime) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    } else {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  } catch (e) {
    Logger.log('Error formatting date: ' + e);
    return "";
  }
}