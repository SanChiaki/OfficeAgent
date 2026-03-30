using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelFocusCoordinatorTests
    {
        [Fact]
        public void RestoreWorksheetFocusActivatesWindowSelectionAndForegroundWindowInOrder()
        {
            var events = new List<string>();
            var coordinator = CreateCoordinator(
                activateActiveWindow: () => events.Add("window"),
                getApplicationWindowHandle: () => new IntPtr(42),
                setActiveWindow: handle =>
                {
                    events.Add("active:" + handle);
                    return handle;
                },
                setFocusWindow: handle =>
                {
                    events.Add("focus:" + handle);
                    return handle;
                },
                setForegroundWindow: handle =>
                {
                    events.Add("foreground:" + handle);
                    return true;
                });

            InvokeRestoreWorksheetFocus(coordinator, () => events.Add("selection"));

            Assert.Equal(
                new[] { "window", "selection", "active:42", "focus:42", "foreground:42" },
                events);
        }

        [Fact]
        public void RestoreWorksheetFocusSkipsForegroundWhenWindowHandleIsZero()
        {
            var events = new List<string>();
            var coordinator = CreateCoordinator(
                activateActiveWindow: () => events.Add("window"),
                getApplicationWindowHandle: () => IntPtr.Zero,
                setActiveWindow: handle =>
                {
                    events.Add("active:" + handle);
                    return handle;
                },
                setFocusWindow: handle =>
                {
                    events.Add("focus:" + handle);
                    return handle;
                },
                setForegroundWindow: handle =>
                {
                    events.Add("foreground:" + handle);
                    return true;
                });

            InvokeRestoreWorksheetFocus(coordinator, () => events.Add("selection"));

            Assert.Equal(
                new[] { "window", "selection" },
                events);
        }

        private static object CreateCoordinator(
            Action activateActiveWindow,
            Func<IntPtr> getApplicationWindowHandle,
            Func<IntPtr, IntPtr> setActiveWindow,
            Func<IntPtr, IntPtr> setFocusWindow,
            Func<IntPtr, bool> setForegroundWindow)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var coordinatorType = addInAssembly.GetType(
                "OfficeAgent.ExcelAddIn.Excel.ExcelFocusCoordinator",
                throwOnError: true);

            return Activator.CreateInstance(
                coordinatorType,
                BindingFlags.Instance | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { activateActiveWindow, getApplicationWindowHandle, setActiveWindow, setFocusWindow, setForegroundWindow },
                culture: null);
        }

        private static void InvokeRestoreWorksheetFocus(object coordinator, Action activateSelection)
        {
            var method = coordinator.GetType().GetMethod(
                "RestoreWorksheetFocus",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            method.Invoke(coordinator, new object[] { activateSelection });
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
