using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
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
    internal static class Utils
    {
        internal static TimeSpan TimeSpanFrom24HString(string str24hour)
        {
            return TimeSpan.ParseExact(str24hour.Replace(":", ""), new string[] { "hhmm", @"hh\:mm" }, CultureInfo.InvariantCulture);
        }

        internal static Process[] FindProcess(string executable)
        {
            return Process.GetProcessesByName(Path.GetFileNameWithoutExtension(executable));
        }

        internal static async Task<bool> DidPredicateBecomeTrueWithinTimeout(Func<bool> predicate, double timeoutInSeconds)
        {
            var cts = new CancellationTokenSource();

            var task = Task.Run(async () =>
            {
                while (!cts.IsCancellationRequested)
                {
                    if (predicate())
                    {
                        return true;
                    }
                    await Task.Delay(5);
                }
                return false;
            });

            if (await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(timeoutInSeconds))) == task)
            {
                return task.Result;
            }
            else
            {
                cts.Cancel();
                return false;
            }
        }

        internal static async Task<IEnumerable<Window>> GetWindowsByClassNameAndProcessNameWithTimeoutAsync(AutomationBase automation, string processName, string classname, double timeoutInSeconds)
        {
            var windows = new List<Window>();
            bool isSuccesful = await DidPredicateBecomeTrueWithinTimeout(() =>
            {
                foreach (var item in FindProcess(processName).Select(x => x.Id))
                {
                    try
                    {
                        foreach (var item2 in Application.Attach(item).GetAllTopLevelWindows(automation))
                        {
                            try
                            {
                                if (item2.ClassName == classname)
                                {
                                    windows.Add(item2);
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                return windows.Count() > 0;
            }, timeoutInSeconds);

            if (isSuccesful)
            {
                return windows;
            }
            else
            {
                throw new Exception("Timed out");
            }
        }

        internal static void ClickButtonInWindowByText(Window window, string text)
        {
            var button = window.FindFirstDescendant(x => x.ByText(text).And(x.ByControlType(FlaUI.Core.Definitions.ControlType.Button)))?.AsButton();
            button.WaitUntilEnabled(TimeSpan.FromSeconds(10));

            DidPredicateBecomeTrueWithinTimeout(() =>
            {
                try
                {
                    button.Click();
                    return true;
                }
                catch (FlaUI.Core.Exceptions.NoClickablePointException) { }
                return false;
            }, 10).Wait();
        }

        internal static void SetEditControlInputByText(Window window, string text, string input)
        {
            var inputBox = window.FindFirstDescendant(x => x.ByText(text))?.AsTextBox();
            inputBox.Focus();
            Keyboard.Type(input);
        }
    }
}
