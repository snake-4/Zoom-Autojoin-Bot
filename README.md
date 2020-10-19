# Yet-Another-Zoom-Autojoin-Bot-CSharp (YAZAB) [![HitCount](http://hits.dwyl.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot-CSharp.svg)](http://hits.dwyl.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot-CSharp)

YAZAB-CSharp is the C# version of [this autojoin bot](https://github.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot) that I've wrote. I'll no longer update the Python version because this version is better.

## Installation
#### Download the [latest release from here](https://github.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot-CSharp/releases)
* **Important**: Make sure your computer has the .NET Framework 4.5
## Usage

#### Create a .csv file for scheduled meetings, take the CSV format below as reference.
```
daysInteger,timeIn24H,meetingId,meetingPsw,meetingTimeInSeconds,comment
1,09:00,111222333,password,3600,Physics
8,12:00,111222333,password,3600,Job interview
```
* `daysInteger` field is a bit-field, bits from 0 to 7 signify the days of the week starting from Monday.
* Other fields are self explanatory.

#### Run the program and enter the path to the .csv file when asked.
* **Important**: Make sure Zoom is installed and is up-to-date!

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://github.com/SnakePin/Yet-Another-Zoom-Autojoin-Bot-CSharp/blob/main/LICENSE)