using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Application = FlaUI.Core.Application;

namespace YAZABNET
{
    static class Utils
    {
        public static TimeSpan TimeSpanFrom24HString(string str24hour)
        {
            str24hour = str24hour.Replace(":", "").PadLeft(4, '0');
            return TimeSpan.ParseExact(str24hour, new string[] { "hhmm", @"hh\:mm" }, CultureInfo.InvariantCulture);
        }

        public static Process[] FindProcess(string executable)
        {
            return Process.GetProcessesByName(Path.GetFileNameWithoutExtension(executable));
        }

        public static IEnumerable<Window> GetTopLevelWindowsByClassName(AutomationBase automation, string className)
        {
            var desktop = automation.GetDesktop();
            var foundElements = desktop.FindAllChildren(cf => cf
                .ByControlType(ControlType.Window)
                .And(cf.ByClassName(className)));
            //This expression right here eats the CPU because of the frequent COM calls

            return foundElements.Select(x => x.AsWindow());
        }

        public static IEnumerable<Window> GetWindowsByClassNameAndProcessNameWithTimeout(AutomationBase automation, string processName, string className, TimeSpan timeout)
        {
            //Once the Retry.While API is asynchronous, I'll make most of these functions async
            //As using Task.Run() would provide no benefits of async as the function blocks the whole thread and each call would spawn a new thread
            //As of now the Retry.While API doesn't support cancellation tokens either, I'm going to add cancellation token stuff too once it does
            return Retry.WhileEmpty(() =>
            {
                var processIDs = FindProcess(processName).Select(x => x.Id);
                return GetTopLevelWindowsByClassName(automation, className).Where(x => processIDs.Contains(x.Properties.ProcessId));
            }, timeout, TimeSpan.FromMilliseconds(500), true, true).Result;
        }

        public static void ClickButtonInWindowByText(Window window, string text)
        {
            var button = (Button)Retry.Find(() => window.FindFirstDescendant(x => x.ByText(text).And(x.ByControlType(ControlType.Button))).AsButton(),
                 new RetrySettings
                 {
                     Timeout = TimeSpan.FromSeconds(10),
                     Interval = TimeSpan.FromMilliseconds(500)
                 }
             );

            button.WaitUntilClickable(TimeSpan.FromSeconds(5))
                .WaitUntilEnabled(TimeSpan.FromSeconds(5))
                .Click();
        }

        public static void SetEditControlInputByText(Window window, string text, string input)
        {
            var inputBox = (TextBox)Retry.Find(() => window.FindFirstDescendant(x => x.ByText(text).And(x.ByControlType(ControlType.Edit))).AsTextBox(),
                 new RetrySettings
                 {
                     Timeout = TimeSpan.FromSeconds(10),
                     Interval = TimeSpan.FromMilliseconds(500)
                 }
             );

            inputBox.WaitUntilClickable(TimeSpan.FromSeconds(5))
                .WaitUntilEnabled(TimeSpan.FromSeconds(5))
                .Focus();
            Keyboard.Type(input);
        }
    }
}
