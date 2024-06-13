# Excel Column Navigator Add-In

## Overview

The Excel Column Navigator Add-In is designed to enhance the user experience in Microsoft Excel by providing advanced column navigation, management, and analysis features. This add-in helps users efficiently navigate large datasets, organize columns, and gain valuable insights through detailed column profiles.

## Features

### Column Navigation
- **Vertical Column Display**: View all columns vertically within a pop-up, making navigation simpler and more intuitive.
- **Find Feature**: Quickly locate columns by name using the search functionality.
- **Multi-Sheet Dropdown**: Seamlessly navigate and manage columns across multiple sheets within a workbook.
- **Column Profile**: Provides detailed statistics for selected columns, including minimum, maximum, and average values.

### Data Organization
- **Sort by Names**: Organize columns alphabetically or by custom criteria.
- **Single Button Toggle for Sorting**: Quickly switch between ascending and descending order.
- **Default Order Implementation**: Maintain a consistent and familiar column order for easier data management.
- **<missing name> Handling**: Identify and highlight columns with missing names to ensure data completeness.

### Data Management
- **Hide/Unhide Functionality**: Simplify data presentation by hiding irrelevant columns and focusing on relevant data.
- **Lock Sheet Button**: Prevent accidental changes to critical data by locking sheets.
- **Auto Refresh**: Automatically update column information based on context changes, ensuring data is always current.

### User Interface
- **Fabric UI Integration**: Adopts Microsoft’s design language for a consistent and professional look.
- **UI/UX Polishing**: Ensures an intuitive, user-friendly interface with minimal learning curve.

## Getting Started

### Cloning the Repository and Running Locally

1. **Clone the Repository**
   ```bash
   git clone https://github.com/tejanshsachdeva/Taskpane.git
   cd GoToColumn
   ```

2. **Install Dependencies**
   ```bash
   npm install
   ```

3. **Initialize the Project**
   ```bash
   npm init
   ```

4. **Run the Add-In Locally**
   ```bash
   npm run start
   ```
   This will start a local server and open Excel with the add-in loaded.

### Sideloading the Add-In

1. **Download the Manifest File**
   - [Manifest File + Documentation](https://drive.google.com/drive/folders/1hInO0tXNOObXB88Kw5Bz1HFcQ7fjUm6g?usp=sharing)

2. **Sideload the Add-In**
   - Open Excel and go to `File -> Options -> Trust Center -> Trust Center Settings -> Trusted Add-in Catalogs`.
   - Add the URL of the folder where `manifest.xml` is located.
   - Go to `Insert -> My Add-ins -> See All... -> SHARED FOLDER` and select your add-in.

Following these steps, you can easily clone the repository and run the add-in locally or sideload the add-in using the manifest file to start using the Excel Column Navigator Add-In.
