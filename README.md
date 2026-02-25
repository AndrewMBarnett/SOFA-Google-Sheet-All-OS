# Apple OS Security Updates Tracker

A Google Apps Script that fetches and displays security update information for all Apple platforms in a Google Sheets dashboard.

## Overview

This script automatically retrieves security update data from the [SOFA (Simple Organized Feed for Apple Software Updates)](https://sofa.macadmins.io/) API and organizes it in a Google Sheets spreadsheet. It tracks updates for macOS, iOS, tvOS, watchOS, visionOS, and Safari, including CVE counts and actively exploited vulnerabilities.

## Features

- **Centralized Dashboard** - View latest updates for all Apple platforms at a glance
- **Security Focus** - Track CVE counts and actively exploited vulnerabilities
- **Color-Coded Interface** - Platform-specific colors with highlighting for exploited CVEs
- **Timestamp Tracking** - Automatic last-updated timestamps on all sheets
- **SOFA Integration** - Direct links to detailed SOFA pages for each OS version
- **Batch Updates** - Update all platforms with a single click

## Installation

1. Open Google Sheets and create a new spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code in the script editor
4. Copy and paste this entire script into the editor
5. Click the **Save** icon and give your project a name
6. Close the Apps Script editor and refresh your Google Sheet
7. You'll see a new menu called **"OS Updates Menu"** appear

## Usage

### Custom Menu Options

After installation, you'll find these options in the **OS Updates Menu**:

#### Individual Platform Updates
- **Fetch macOS Data** - Updates only macOS sheet
- **Fetch iOS Data** - Updates only iOS sheet
- **Fetch tvOS Data** - Updates only tvOS sheet
- **Fetch watchOS Data** - Updates only watchOS sheet
- **Fetch visionOS Data** - Updates only visionOS sheet
- **Fetch Safari Data** - Updates only Safari sheet

#### Bulk Operations
- **Fetch All Platforms** - Updates all platform sheets and the dashboard
- **Update Dashboard Only** - Refreshes the dashboard from existing data

### First Time Setup

1. Click **OS Updates Menu > Fetch All Platforms**
2. Grant necessary permissions when prompted
3. Wait for the script to complete (you'll see a success message)

### Sheet Structure

The script creates the following sheets:

#### Dashboard Sheet
Displays the latest version and security release for each platform with:
- Platform name
- OS/App version
- Latest update name
- Product version
- Release date
- CVE count
- Actively exploited CVEs (highlighted in red)
- Days since previous release
- Security info URL
- SOFA details link

#### Individual Platform Sheets
Each platform (macOS, iOS, tvOS, watchOS, visionOS, Safari) has a detailed sheet with:
- Complete version history
- All security releases
- Detailed CVE information
- Release dates and intervals

## Data Sources

All data is fetched from the official SOFA feeds:
- macOS: `https://sofa.macadmins.io/v2/macos_data_feed.json`
- iOS: `https://sofafeed.macadmins.io/v2/ios_data_feed.json`
- tvOS: `https://sofa.macadmins.io/v2/tvos_data_feed.json`
- watchOS: `https://sofa.macadmins.io/v2/watchos_data_feed.json`
- visionOS: `https://sofa.macadmins.io/v2/visionos_data_feed.json`
- Safari: `https://sofa.macadmins.io/v2/safari_data_feed.json`

## Color Coding

The dashboard uses platform-specific colors for easy identification:
- **macOS**: Blue (`#E8F0FE`)
- **iOS**: Orange (`#FFF3E0`)
- **tvOS**: Purple (`#F3E5F5`)
- **watchOS**: Teal (`#E0F2F1`)
- **visionOS**: Pink (`#FCE4EC`)
- **Safari**: Light Blue (`#E3F2FD`)

**Actively exploited CVEs are highlighted in red** (`#F4CCCC`) with bold text.

## Automation

To set up automatic updates:

1. In the Apps Script editor, click the **clock icon** for Triggers
2. Click **+ Add Trigger**
3. Configure:
   - Function: `fetchAllPlatforms`
   - Event source: Time-driven
   - Type: Choose your preferred schedule (e.g., "Day timer")
   - Time of day: Select preferred time
4. Click **Save**

Recommended schedule: Daily updates at a time that works for your team.

## Troubleshooting

### Permission Errors
- Make sure you've authorized the script to access external URLs
- Check that the spreadsheet isn't in restricted sharing mode

### Data Not Appearing
- Check the Apps Script logs: **Extensions > Apps Script > Executions**
- Verify internet connectivity
- Ensure SOFA API endpoints are accessible

### Slow Performance
- The script processes multiple API calls; expect 30-60 seconds for full updates
- Consider updating individual platforms instead of all at once during testing

## Credits

- **Data Source**: [SOFA by MacAdmins](https://sofa.macadmins.io/)

## Version Support

The script currently includes SOFA links for:
- macOS: Monterey (12) through Tahoe (26)
- iOS: iOS 17, 18, 26
- Safari: Safari 17, 18, 26
- tvOS: tvOS 18, 26
- visionOS: visionOS 2, 26
- watchOS: watchOS 11, 26

## License

This script is provided as-is for use by the MacAdmin community.

---

**Note**: This tool is for tracking security updates. Always refer to official Apple security documentation for production deployment decisions.