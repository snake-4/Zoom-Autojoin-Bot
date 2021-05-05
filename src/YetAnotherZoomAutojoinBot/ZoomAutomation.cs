using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Tools;
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
    class ZoomAutomation
    {
        const string zoomExecutableName = "Zoom.exe";
        const string ZPFTEWndClass_LoadingConnectingString = "Connecting\u2026"; //Both signed-in and anonymous users loading screen
        const string ZPFTEWndClass_JoinAMeetingBtnString = "Join a Meeting"; //Anonymous users menu
        const string ZPPTMainFrmWndClassEx_JoinBtnString = "Join"; //Signed-in users menu
        const string zWaitHostWndClass_JoinBtnString = "Join"; //Actual join meeting menu
        const string zWaitHostWndClass_MeetingIDTextBoxString = "Please enter your Meeting ID or Personal Link Name";
        const string zWaitHostWndClass_UserNameTextBoxString = "Please enter your name";
        const string zWaitHostWndClass_MeetingPasscodeScreenTitle = "Enter meeting passcode";
        const string VideoPreviewWndClass_JoinWithoutVideoBtnString = "Join without Video";
        const string zWaitHostWndClass_PasswordTextBoxString = "Please enter meeting passcode";
        const string zWaitHostWndClass_PasswordScreenJoinBtnString = "Join Meeting";

        enum ZoomState
        {
            InMeeting,
            CamPreviewWindow,
            JoiningFailed,
            PasswordRequired
        };

        AutomationBase automation = new UIA3Automation();
        TimeSpan defaultIntervalForFunctions = TimeSpan.FromMilliseconds(500);

        string GetZoomPath()
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

        void KillZoom()
        {
            foreach (var item in Utils.FindProcess(zoomExecutableName))
            {
                item.Kill();
            }
        }

        void StartZoom()
        {
            KillZoom();
            Process.Start(GetZoomPath()).WaitForInputIdle();
        }

        async Task<List<Window>> GetZoomWindowsByClassNameWithTimeoutAsync(string classname, TimeSpan timeout)
        {
            return await Utils.GetWindowsByClassNameAndProcessNameWithTimeoutAsync(automation, zoomExecutableName, classname, timeout);
        }

        Window GetZoomMainMenu(TimeSpan timeout)
        {
            var signedInMenu = GetZoomWindowsByClassNameWithTimeoutAsync("ZPPTMainFrmWndClassEx", timeout);
            var anonymousMenu = GetZoomWindowsByClassNameWithTimeoutAsync("ZPFTEWndClass", timeout);

            while (!signedInMenu.IsCompleted || !anonymousMenu.IsCompleted)
            {
                Thread.Sleep(defaultIntervalForFunctions);

                if (signedInMenu.IsCompleted && !signedInMenu.IsFaulted)
                {
                    if (signedInMenu.Result.Count > 0)
                    {
                        return signedInMenu.Result.First();
                    }
                }

                if (anonymousMenu.IsCompleted && !anonymousMenu.IsFaulted)
                {
                    if (anonymousMenu.Result.Count > 0)
                    {
                        return anonymousMenu.Result.First();
                    }
                }
            }

            throw new Exception("Main menu couldn't be found.");
        }

        void OpenZoomJoinMenu(TimeSpan timeout)
        {
            Retry.WhileFalse(() =>
            {
                var menu = GetZoomMainMenu(timeout);
                if (menu.ClassName == "ZPPTMainFrmWndClassEx")
                {
                    Utils.ClickButtonInWindowByText(menu, ZPPTMainFrmWndClassEx_JoinBtnString);
                }
                else if (menu.ClassName == "ZPFTEWndClass")
                {
                    if (menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_LoadingConnectingString)).Length > 0
                        || menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_JoinAMeetingBtnString)).Length == 0)
                    {
                        return false;
                    }
                    Utils.ClickButtonInWindowByText(menu, ZPFTEWndClass_JoinAMeetingBtnString);
                }

                return true;
            }, timeout, defaultIntervalForFunctions, true, true);
        }

        void ZoomEnterIDAndJoin(TimeSpan timeout, string meetingid, string username = null)
        {
            var joinMenu = GetZoomWindowsByClassNameWithTimeoutAsync("zWaitHostWndClass", timeout).Result.First();

            if (username != null)
            {
                Utils.SetEditControlInputByText(joinMenu, zWaitHostWndClass_UserNameTextBoxString, username);
            }
            Utils.SetEditControlInputByText(joinMenu, zWaitHostWndClass_MeetingIDTextBoxString, meetingid);
            Utils.ClickButtonInWindowByText(joinMenu, zWaitHostWndClass_JoinBtnString);
        }

        async Task<Window> GetZoomPasswordScreenWindow(TimeSpan timeout)
        {
            Window passwordMenu = null;
            if (!await Utils.DidPredicateBecomeTrueWithinTimeout(() =>
             {
                 var windows = GetZoomWindowsByClassNameWithTimeoutAsync("zWaitHostWndClass", timeout).Result;
                 foreach (var window in windows)
                 {
                     if (window.IsAvailable && window.Title == zWaitHostWndClass_MeetingPasscodeScreenTitle)
                     {
                         passwordMenu = window;
                         return true;
                     }
                 }
                 return false;
             }, timeout))
            {
                throw new Exception("Timed out");
            }
            return passwordMenu;
        }

        ZoomState GetZoomStateAfterJoin(TimeSpan timeout)
        {
            var findFailedTask = GetZoomWindowsByClassNameWithTimeoutAsync("zJoinMeetingFailedDlgClass", timeout);
            var findCamPreviewTask = GetZoomWindowsByClassNameWithTimeoutAsync("VideoPreviewWndClass", timeout);
            var mainWindowTask = GetZoomWindowsByClassNameWithTimeoutAsync("ZPContentViewWndClass", timeout);
            var passwordRequiredTask = GetZoomPasswordScreenWindow(timeout);

            do
            {
                if (findCamPreviewTask.IsCompleted && !findCamPreviewTask.IsFaulted)
                {
                    return ZoomState.CamPreviewWindow;
                }
                else if (mainWindowTask.IsCompleted && !mainWindowTask.IsFaulted)
                {
                    return ZoomState.InMeeting;
                }
                else if (findFailedTask.IsCompleted && !findFailedTask.IsFaulted)
                {
                    return ZoomState.JoiningFailed;
                }
                else if (passwordRequiredTask.IsCompleted && !passwordRequiredTask.IsFaulted)
                {
                    return ZoomState.PasswordRequired;
                }
                Thread.Sleep(500);
            } while (!findFailedTask.IsCompleted || !findCamPreviewTask.IsCompleted || !mainWindowTask.IsCompleted || !passwordRequiredTask.IsCompleted);

            return ZoomState.JoiningFailed;
        }

        void ZoomWaitUntilStateIsNot(ZoomState stateNotToBe, TimeSpan timeout)
        {
            _ = Retry.WhileFalse(() =>
            {
                return GetZoomStateAfterJoin(timeout) != stateNotToBe;
            }, timeout, defaultIntervalForFunctions, true);
        }

        void ZoomSkipCameraDialog(TimeSpan timeout)
        {
            _ = Retry.WhileException(() =>
            {
                Window findCamPreviewWindow = GetZoomWindowsByClassNameWithTimeoutAsync("VideoPreviewWndClass", timeout).Result.First();
                Utils.ClickButtonInWindowByText(findCamPreviewWindow, VideoPreviewWndClass_JoinWithoutVideoBtnString);
            }, timeout, defaultIntervalForFunctions, true);
        }

        void ZoomEnterPasswordAndJoin(string meetingPSW, TimeSpan timeout)
        {
            var pswMenu = GetZoomPasswordScreenWindow(timeout).Result;
            Utils.SetEditControlInputByText(pswMenu, zWaitHostWndClass_PasswordTextBoxString, meetingPSW);
            Utils.ClickButtonInWindowByText(pswMenu, zWaitHostWndClass_PasswordScreenJoinBtnString);
        }

        bool JoinZoom(string meetingID, string meetingPSW)
        {
            TimeSpan timeoutForGUIFunctions = TimeSpan.FromSeconds(15);
            TimeSpan timeoutForZoomStart = TimeSpan.FromSeconds(30);

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
                OpenZoomJoinMenu(timeoutForZoomStart);
            }
            catch (Exception exc)
            {
                Console.WriteLine("Main menu couldn't be found or interacted with. Exception: " + exc);
            }

            ZoomEnterIDAndJoin(timeoutForGUIFunctions, meetingID);
            if (GetZoomStateAfterJoin(timeoutForGUIFunctions) == ZoomState.PasswordRequired)
            {
                ZoomEnterPasswordAndJoin(meetingPSW, timeoutForGUIFunctions);
                ZoomWaitUntilStateIsNot(ZoomState.PasswordRequired, timeoutForGUIFunctions);
            }

            if (GetZoomStateAfterJoin(timeoutForGUIFunctions) == ZoomState.JoiningFailed)
            {
                return false;
            }

            if (GetZoomStateAfterJoin(timeoutForGUIFunctions) == ZoomState.CamPreviewWindow)
            {
                ZoomSkipCameraDialog(timeoutForGUIFunctions);
                ZoomWaitUntilStateIsNot(ZoomState.CamPreviewWindow, timeoutForGUIFunctions);
            }
            return true;
        }

        public void ZoomJoinWaitLeave(string meetingID, string meetingPSW, int meetingTimeInSeconds)
        {
            JoinZoom(meetingID, meetingPSW);
            Thread.Sleep(meetingTimeInSeconds * 1000);
            KillZoom();
        }
    }
}
