# Yet Another Zoom Autojoin Bot (C#)

This is the C# version of [this autojoin bot](https://github.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot) that I've wrote. I'll no longer update the Python version because this version is better.

## Installation
#### 1. Download the [latest release from here](https://github.com/SnakePin/Zoom-Autojoin-Bot/releases)
* **Important**: Make sure your computer has the .NET Framework 4.5
## Usage

#### 1. Create a .csv file for scheduled meetings, take the CSV format below as reference.
```
dayOfWeek,timeIn24H,meetingId,meetingPsw,meetingTimeInSeconds,comment
1,09:00,111222333,password,3600,Physics
2,12:00,111222333,password,3600,Job interview
```
* `dayOfWeek` field can have the value of any number from 1 to 7 with each number representing a day of the week starting from Monday.
* Other fields are self explanatory.

#### 2. Run the program and enter the path to the .csv file when asked.
* **Important**: Make sure Zoom is installed and is up-to-date!
* **Important**: Your Zoom client's language **MUST** be set to **English**.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://github.com/SnakePin/Zoom-Autojoin-Bot/blob/main/LICENSE)
