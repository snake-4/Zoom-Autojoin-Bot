# Yet Another Zoom Autojoin Bot (C#)

This is a Zoom auto-join bot written in C#. It can be configured using a CSV file.

## Installation
#### 1. Download the [latest release from here](https://github.com/SnakePin/Zoom-Autojoin-Bot/releases)
* **Important**: Make sure your computer has the .NET Framework 4.5
## Usage

#### 1. Create a .csv file for scheduled meetings, take the CSV format below as reference.
```
DayOfWeek,MeetingTimeIn24H,MeetingID,MeetingPassword,MeetingTimeInSeconds,Comment
1,09:00,111222333,password,3600,Physics
2,12:00,111222333,password,3600,Job interview
```
* `DayOfWeek` field can have the value of any number from 1 to 7 with each number representing a day of the week starting from Monday.
* Other fields are self explanatory.

#### 2. Run the program and enter the path to the .csv file when asked.
* **Important**: Make sure Zoom is installed and is up-to-date!
* **Important**: Your Zoom client's language **MUST** be set to **English**.

#### 2.1 Running the program from command line
The path of the CSV file can also be passed as an argument to the program like this `program.exe "PathToCsv"`
thus you can create a batch file to start the program without it asking the CSV path.

#### 2.2 Drag-and-drop the .csv file
You can also just drag-and-drop the .csv file on the program's executable.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://github.com/SnakePin/Zoom-Autojoin-Bot/blob/main/LICENSE)
