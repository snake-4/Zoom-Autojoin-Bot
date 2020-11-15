using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace YAZABNET
{
    internal static class ZoomAutomationFunctions
    {
        const string ZPFTEWndClass_LoadingConnectingString = "Connecting\u2026";
        const string ZPFTEWndClass_JoinAMeetingBtnString = "Join a Meeting";
        const string zoomExecutableName = "Zoom.exe";

        static string GetZoomPath()
        {
            var key1 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ZoomUMX", "InstallLocation", null);
            if (key1 != null)
            {
                return Path.Combine(key1, zoomExecutableName);
            }

            var key2 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ZoomUMX", "UninstallString", null);
            if (key2 != null)
            {
                return Path.Combine(key2.Replace("\"", "").Replace("\\uninstall\\Installer.exe /uninstall", ""), zoomExecutableName);
            }

            var key3 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Classes\\ZoomLauncher\\shell\\open\\command", "(Default)", null);
            if (key3 != null)
            {
                return Path.Combine(key3.Replace("\"", "").Replace("--url=%1", ""), zoomExecutableName);
            }

            return null;
        }

        static void KillZoom()
        {
            foreach (var item in Utils.FindProcess(zoomExecutableName))
            {
                item.Kill();
            }
        }

        static void StartZoom()
        {
            KillZoom();
            Process.Start(GetZoomPath()).WaitForInputIdle();
        }

        static async Task<IEnumerable<Window>> GetZoomWindowsByClassNameWithTimeoutAsync(AutomationBase automation, string classname, double timeoutInSeconds)
        {
            return await Utils.GetWindowsByClassNameAndProcessNameWithTimeoutAsync(automation, zoomExecutableName, classname, timeoutInSeconds);
        }

        static Window GetZoomMainMenu(AutomationBase automation, int timeoutInSeconds)
        {
            var signedInMenu = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "ZPPTMainFrmWndClassEx", timeoutInSeconds);
            var anonymousMenu = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "ZPFTEWndClass", timeoutInSeconds);

            while (!signedInMenu.IsCompleted || !anonymousMenu.IsCompleted)
            {
                Thread.Sleep(10);

                if (signedInMenu.IsCompleted && !signedInMenu.IsFaulted)
                {
                    if (signedInMenu.Result.Count() > 0)
                    {
                        return signedInMenu.Result.First();
                    }
                }

                if (anonymousMenu.IsCompleted && !anonymousMenu.IsFaulted)
                {
                    if (anonymousMenu.Result.Count() > 0)
                    {
                        return anonymousMenu.Result.First();
                    }
                }
            }

            throw new Exception("Main menu couldn't be found.");
        }

        static void OpenZoomJoinMenu(AutomationBase automation)
        {
        TRY_AGAIN:
            var menu = GetZoomMainMenu(automation, 30);
            if (menu.ClassName == "ZPPTMainFrmWndClassEx")
            {
                Utils.ClickButtonInWindowByText(menu, "Join");
            }
            else if (menu.ClassName == "ZPFTEWndClass")
            {
                if (menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_LoadingConnectingString)).Length > 0
                    || menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_JoinAMeetingBtnString)).Length == 0)
                {
                    Thread.Sleep(2000);
                    //Loading screen for signed-in menu
                    goto TRY_AGAIN;
                }
                Utils.ClickButtonInWindowByText(menu, ZPFTEWndClass_JoinAMeetingBtnString);
            }
        }

        static void zoomEnterIdAndPassword(AutomationBase automation, string meetingid, string meetingpsw, string username = null, int timeoutInSeconds = 30)
        {
            var joinMenu = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "zWaitHostWndClass", timeoutInSeconds).Result.First();

            Utils.SetEditControlInputByText(joinMenu, "Please enter your Meeting ID or Personal Link Name", meetingid);

            if (username != null)
            {
                Utils.SetEditControlInputByText(joinMenu, "Please enter your name", username);
            }

            Utils.ClickButtonInWindowByText(joinMenu, "Join");

            if (Utils.DidPredicateBecomeTrueWithinTimeout(() =>
            {
                try
                {
                    joinMenu = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "zWaitHostWndClass", timeoutInSeconds).Result.First();
                    return joinMenu.IsAvailable && joinMenu.Title == "Enter meeting passcode";
                }
                catch
                {
                    return false;
                }
            }, 10).Result)
            {
                Utils.SetEditControlInputByText(joinMenu, "Please enter meeting passcode", meetingpsw);
                Utils.ClickButtonInWindowByText(joinMenu, "Join Meeting");
            }
        }

        static bool zoomFailCheck(AutomationBase automation, int timeoutInSeconds = 5)
        {
            var findFailedTask = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "zJoinMeetingFailedDlgClass", timeoutInSeconds);
            var findCamPreviewTask = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "VideoPreviewWndClass", timeoutInSeconds);
            var mainWindowTask = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "ZPContentViewWndClass", timeoutInSeconds);

            do
            {
                bool camPreviewFound = findCamPreviewTask.IsCompleted && !findCamPreviewTask.IsFaulted;
                bool meetingMainWindowFound = mainWindowTask.IsCompleted && !mainWindowTask.IsFaulted;
                bool failedWindowFound = findFailedTask.IsCompleted && !findFailedTask.IsFaulted;

                if (camPreviewFound || meetingMainWindowFound)
                {
                    return false;
                }
                else if (failedWindowFound)
                {
                    return true;
                }
            } while (!findFailedTask.IsCompleted || !findCamPreviewTask.IsCompleted || !mainWindowTask.IsCompleted);

            return false;
        }

        static void zoomCameraDialogCheck(AutomationBase automation, int timeoutInSeconds = 5)
        {
            Window findCamPreviewWindow;
            try
            {
                findCamPreviewWindow = GetZoomWindowsByClassNameWithTimeoutAsync(automation, "VideoPreviewWndClass", timeoutInSeconds).Result.First();
            }
            catch
            {
                return;
            }
            Utils.ClickButtonInWindowByText(findCamPreviewWindow, "Join without Video");
        }

        static bool JoinZoom(string meetingID, string meetingPSW)
        {
            using (var automation = new UIA3Automation())
            {
                try
                {
                    StartZoom();
                }
                catch (Exception exc)
                {
                    Console.WriteLine("Error while trying to start Zoom. Exception: " + exc);
                }

                try
                {
                    OpenZoomJoinMenu(automation);
                }
                catch (Exception exc)
                {
                    Console.WriteLine("Main menu couldn't be found or interacted with. Exception: " + exc);
                }

                zoomEnterIdAndPassword(automation, meetingID, meetingPSW);
                if (zoomFailCheck(automation))
                {
                    return false;
                }

                zoomCameraDialogCheck(automation);
                return true;
            }
        }

        internal static void ZoomJoinLeaveTask(string meetingID, string meetingPSW, int meetingTimeInSeconds)
        {
            JoinZoom(meetingID, meetingPSW);
            Thread.Sleep(meetingTimeInSeconds * 1000);
            KillZoom();
        }
    }
}
