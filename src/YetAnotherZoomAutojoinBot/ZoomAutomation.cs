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
        const string ZoomExecutableName = "Zoom.exe";
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
        const string AnonymousUserMenuClassName = "ZPFTEWndClass";
        const string MultiplePurposeWindowClassName = "zWaitHostWndClass"; //This window is used for the "waiting for host" window, the join dialog and the passcode dialog.
        const string SignedInUserMenuClassName = "ZPPTMainFrmWndClassEx";
        const string CameraPreviewWindowClassName = "VideoPreviewWndClass";
        const string FailedJoiningMeetingWindowClassName = "zJoinMeetingFailedDlgClass";
        const string InMeetingMenuClassName = "ZPContentViewWndClass";

        enum ZoomState
        {
            JoinedMeeting,
            CamPreviewWindow,
            JoiningFailed,
            PasswordRequired,
            UnknownState
        };

        private AutomationBase AutomationInstance = new UIA3Automation();
        private TimeSpan DefaultIntervalForFunctions = TimeSpan.FromMilliseconds(500);
        private string PathOfZoomExecutable = null;

        public string GetZoomExecutablePath()
        {
            if (PathOfZoomExecutable != null)
            {
                return PathOfZoomExecutable;
            }
            else
            {
                var key1 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ZoomUMX", "InstallLocation", null);
                if (key1 != null)
                {
                    return PathOfZoomExecutable = Path.Combine(key1, ZoomExecutableName);
                }

                var key2 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ZoomUMX", "UninstallString", null);
                if (key2 != null)
                {
                    return PathOfZoomExecutable = Path.Combine(key2.Replace("\"", "").Replace("\\uninstall\\Installer.exe /uninstall", ""), ZoomExecutableName);
                }

                var key3 = (string)Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Classes\\ZoomLauncher\\shell\\open\\command", "(Default)", null);
                if (key3 != null)
                {
                    return PathOfZoomExecutable = Path.Combine(key3.Replace("\"", "").Replace("--url=%1", ""), ZoomExecutableName);
                }
            }

            throw new Exception("Couldn't get the Zoom executable's location. Please make sure that Zoom is installed!");
        }

        public void KillZoom()
        {
            foreach (var item in Utils.FindProcess(ZoomExecutableName))
            {
                item.Kill();
            }
        }

        void StartZoom()
        {
            KillZoom();
            Process.Start(GetZoomExecutablePath()).WaitForInputIdle();
        }

        IEnumerable<Window> GetZoomWindowsByClassNameWithTimeout(string classname, TimeSpan timeout)
           => Utils.GetWindowsByClassNameAndProcessNameWithTimeout(AutomationInstance, ZoomExecutableName, classname, timeout);

        Window GetZoomWindowByClassNameAndTitleWithTimeout(string classname, string title, TimeSpan timeout)
            => Retry.WhileNull(() =>
            {
                return GetZoomWindowsByClassNameWithTimeout(classname, timeout).FirstOrDefault((x) => x.IsAvailable && x.Title == title);
            }, timeout, DefaultIntervalForFunctions, true, true).Result;

        Window GetZoomWaitingForHostWindow(TimeSpan timeout)
           => GetZoomWindowByClassNameAndTitleWithTimeout(MultiplePurposeWindowClassName, zWaitHostWndClass_WaitingForHostScreenTitle, timeout);

        Window GetZoomPasswordScreenWindow(TimeSpan timeout)
           => GetZoomWindowByClassNameAndTitleWithTimeout(MultiplePurposeWindowClassName, zWaitHostWndClass_MeetingPasscodeScreenTitle, timeout);

        Window GetZoomMainMenu(TimeSpan timeout)
        {
            var signedInUserMenu = Task.Run(() => GetZoomWindowsByClassNameWithTimeout(SignedInUserMenuClassName, timeout));
            var anonymousUserMenu = Task.Run(() => GetZoomWindowsByClassNameWithTimeout(AnonymousUserMenuClassName, timeout));

            while (!signedInUserMenu.IsCompleted || !anonymousUserMenu.IsCompleted)
            {
                Thread.Sleep(DefaultIntervalForFunctions);

                if (signedInUserMenu.IsCompleted && !signedInUserMenu.IsFaulted)
                {
                    if (signedInUserMenu.Result.Count() > 0)
                    {
                        return signedInUserMenu.Result.First();
                    }
                }

                if (anonymousUserMenu.IsCompleted && !anonymousUserMenu.IsFaulted)
                {
                    if (anonymousUserMenu.Result.Count() > 0)
                    {
                        return anonymousUserMenu.Result.First();
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
                if (menu.ClassName == SignedInUserMenuClassName)
                {
                    Utils.ClickButtonInWindowByText(menu, ZPPTMainFrmWndClassEx_JoinBtnString);
                }
                else if (menu.ClassName == AnonymousUserMenuClassName)
                {
                    if (menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_LoadingConnectingString)).Length > 0
                        || menu.FindAllDescendants(x => x.ByName(ZPFTEWndClass_JoinAMeetingBtnString)).Length == 0)
                    {
                        return false;
                    }
                    Utils.ClickButtonInWindowByText(menu, ZPFTEWndClass_JoinAMeetingBtnString);
                }

                return true;
            }, timeout, DefaultIntervalForFunctions, true, true);
        }

        void ZoomEnterIDAndJoin(TimeSpan timeout, string meetingid, string username = null)
        {
            var joinMenu = GetZoomWindowsByClassNameWithTimeout(MultiplePurposeWindowClassName, timeout).First();

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
            var task_joiningFailed = Task.Run(() => GetZoomWindowsByClassNameWithTimeout(FailedJoiningMeetingWindowClassName, timeout));
            var task_cameraPreview = Task.Run(() => GetZoomWindowsByClassNameWithTimeout(CameraPreviewWindowClassName, timeout));
            var task_meetingWindow = Task.Run(() => GetZoomWindowsByClassNameWithTimeout(InMeetingMenuClassName, timeout));

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
            timeout, DefaultIntervalForFunctions, true, true).Result;
        }

        void ZoomSkipCameraDialog(TimeSpan timeout)
        {
            _ = Retry.WhileException(() =>
            {
                Window findCamPreviewWindow = GetZoomWindowsByClassNameWithTimeout(CameraPreviewWindowClassName, timeout).First();
                Utils.ClickButtonInWindowByText(findCamPreviewWindow, VideoPreviewWndClass_JoinWithoutVideoBtnString);
            }, timeout, DefaultIntervalForFunctions, true);
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

            StartZoom();
            OpenZoomJoinMenu(timeoutForZoomStart);
            ZoomEnterIDAndJoin(timeoutForGUIFunctions, meetingID);

            ZoomState currentState = GetZoomStateAfterJoin(timeoutForGUIFunctions);

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
