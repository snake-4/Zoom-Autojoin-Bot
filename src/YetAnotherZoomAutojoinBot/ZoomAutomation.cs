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
        const string zWaitHostWndClass_WaitingForHostScreenTitle = "Waiting for Host";
        const string VideoPreviewWndClass_JoinWithoutVideoBtnString = "Join without Video";
        const string zWaitHostWndClass_PasswordTextBoxString = "Please enter meeting passcode";
        const string zWaitHostWndClass_PasswordScreenJoinBtnString = "Join Meeting";

        enum ZoomState
        {
            JoinedMeeting,
            CamPreviewWindow,
            JoiningFailed,
            PasswordRequired,
            UnknownState
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

        public void KillZoom()
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

        IEnumerable<Window> GetZoomWindowsByClassNameWithTimeout(string classname, TimeSpan timeout)
           => Utils.GetWindowsByClassNameAndProcessNameWithTimeout(automation, zoomExecutableName, classname, timeout);

        Window GetZoomWindowByClassNameAndTitleWithTimeout(string classname, string title, TimeSpan timeout)
            => Retry.WhileNull(() =>
            {
                return GetZoomWindowsByClassNameWithTimeout(classname, timeout).FirstOrDefault((x) => x.IsAvailable && x.Title == title);
            }, timeout, defaultIntervalForFunctions, true, true).Result;

        Window GetZoomWaitingForHostWindow(TimeSpan timeout)
           => GetZoomWindowByClassNameAndTitleWithTimeout("zWaitHostWndClass", zWaitHostWndClass_WaitingForHostScreenTitle, timeout);

        Window GetZoomPasswordScreenWindow(TimeSpan timeout)
           => GetZoomWindowByClassNameAndTitleWithTimeout("zWaitHostWndClass", zWaitHostWndClass_MeetingPasscodeScreenTitle, timeout);

        Window GetZoomMainMenu(TimeSpan timeout)
        {
            var signedInMenu = Task.Run(() => GetZoomWindowsByClassNameWithTimeout("ZPPTMainFrmWndClassEx", timeout));
            var anonymousMenu = Task.Run(() => GetZoomWindowsByClassNameWithTimeout("ZPFTEWndClass", timeout));

            while (!signedInMenu.IsCompleted || !anonymousMenu.IsCompleted)
            {
                Thread.Sleep(defaultIntervalForFunctions);

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
            var joinMenu = GetZoomWindowsByClassNameWithTimeout("zWaitHostWndClass", timeout).First();

            if (username != null)
            {
                Utils.SetEditControlInputByText(joinMenu, zWaitHostWndClass_UserNameTextBoxString, username);
            }
            Utils.SetEditControlInputByText(joinMenu, zWaitHostWndClass_MeetingIDTextBoxString, meetingid);
            Utils.ClickButtonInWindowByText(joinMenu, zWaitHostWndClass_JoinBtnString);
        }

        ZoomState GetZoomStateAfterJoin(TimeSpan timeout)
        {
            var task_passwordRequired = Task.Run(() => GetZoomPasswordScreenWindow(timeout));
            var task_waitingForHost = Task.Run(() => GetZoomWaitingForHostWindow(timeout));
            var task_joiningFailed = Task.Run(() => GetZoomWindowsByClassNameWithTimeout("zJoinMeetingFailedDlgClass", timeout));
            var task_cameraPreview = Task.Run(() => GetZoomWindowsByClassNameWithTimeout("VideoPreviewWndClass", timeout));
            var task_meetingWindow = Task.Run(() => GetZoomWindowsByClassNameWithTimeout("ZPContentViewWndClass", timeout));

            //TODO: cancel all the other tasks when one of them completes

            do
            {
                if (task_cameraPreview.IsCompleted && !task_cameraPreview.IsFaulted)
                {
                    return ZoomState.CamPreviewWindow;
                }
                else if ((task_meetingWindow.IsCompleted && !task_meetingWindow.IsFaulted)
                    || (task_waitingForHost.IsCompleted && !task_waitingForHost.IsFaulted))
                {
                    return ZoomState.JoinedMeeting;
                }
                else if (task_joiningFailed.IsCompleted && !task_joiningFailed.IsFaulted)
                {
                    return ZoomState.JoiningFailed;
                }
                else if (task_passwordRequired.IsCompleted && !task_passwordRequired.IsFaulted)
                {
                    return ZoomState.PasswordRequired;
                }

                Thread.Sleep(500);
            } while (!task_joiningFailed.IsCompleted
            || !task_cameraPreview.IsCompleted
            || !task_meetingWindow.IsCompleted
            || !task_passwordRequired.IsCompleted
            || !task_waitingForHost.IsCompleted);

            //throw new Exception("Zoom after join state cannot be determined.");
            //Throw TimeoutException here instead?
            return ZoomState.UnknownState;
        }

        //Returns the new state, will throw on timeout
        ZoomState ZoomWaitUntilStateIsNot(ZoomState stateNotToBe, TimeSpan timeout)
        {
            return Retry.While(() =>
            {
                return GetZoomStateAfterJoin(timeout);
            }, (state) => state == stateNotToBe,
            timeout, defaultIntervalForFunctions, true, true).Result;
        }

        void ZoomSkipCameraDialog(TimeSpan timeout)
        {
            _ = Retry.WhileException(() =>
            {
                Window findCamPreviewWindow = GetZoomWindowsByClassNameWithTimeout("VideoPreviewWndClass", timeout).First();
                Utils.ClickButtonInWindowByText(findCamPreviewWindow, VideoPreviewWndClass_JoinWithoutVideoBtnString);
            }, timeout, defaultIntervalForFunctions, true);
        }

        void ZoomEnterPasswordAndJoin(string meetingPSW, TimeSpan timeout)
        {
            var pswMenu = GetZoomPasswordScreenWindow(timeout);
            Utils.SetEditControlInputByText(pswMenu, zWaitHostWndClass_PasswordTextBoxString, meetingPSW);
            Utils.ClickButtonInWindowByText(pswMenu, zWaitHostWndClass_PasswordScreenJoinBtnString);
        }

        public bool JoinZoom(string meetingID, string meetingPSW)
        {
            TimeSpan timeoutForGUIFunctions = TimeSpan.FromSeconds(15);
            TimeSpan timeoutForZoomStart = TimeSpan.FromSeconds(30);
            ZoomState currentState = ZoomState.UnknownState;

            StartZoom();

            OpenZoomJoinMenu(timeoutForZoomStart);

            ZoomEnterIDAndJoin(timeoutForGUIFunctions, meetingID);
            currentState = GetZoomStateAfterJoin(timeoutForGUIFunctions);

            if (currentState == ZoomState.PasswordRequired)
            {
                ZoomEnterPasswordAndJoin(meetingPSW, timeoutForGUIFunctions);
                currentState = ZoomWaitUntilStateIsNot(ZoomState.PasswordRequired, timeoutForGUIFunctions);
            }

            if (currentState == ZoomState.CamPreviewWindow)
            {
                ZoomSkipCameraDialog(timeoutForGUIFunctions);
                currentState = ZoomWaitUntilStateIsNot(ZoomState.CamPreviewWindow, timeoutForGUIFunctions);
            }

            return currentState == ZoomState.JoinedMeeting;
        }
    }
}
