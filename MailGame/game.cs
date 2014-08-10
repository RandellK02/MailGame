using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;

namespace MailGame
{
    internal class game
    {
        private static string victim = "Reese";
        private static int sleepCount = 100;
        private static int longSleepCount = 3000;

        private static void Main(string[] args)
        {
            //Intro();
            play();
        }

        private static void play()
        {
            AutomationElement mainWnd;

            getOutlookWindow(ref mainWnd);
            openNewEmail(mainWnd);
        }

        private static void openNewEmail(AutomationElement mainWnd)
        {
            PropertyCondition idCond;
            AutomationElement element;
            idCond = new PropertyCondition(AutomationElement.AutomationIdProperty, NextBtnID);
            element = mainWnd.FindFirst(TreeScope.Descendants, idCond);
            while (element == null)
            {
                element = mainWnd.FindFirst(TreeScope.Descendants, idCond);
            }

            element.SetFocus();
            //Get invoke pattern and invoke it to press the button
            InvokePattern invPattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invPattern.Invoke();
        }

        private static void getOutlookWindow(ref AutomationElement mainWnd)
        {
            //Create a property condition with the element's type
            PropertyCondition typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            //Create a property condition with the element's name
            PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "Inbox - Randell.Koen@CalRecycle.ca.gov - Outlook");
            //Create the conjunction condition
            AndCondition andCondition = new AndCondition(typeCondition, nameCondition);
            //Ask the Desktop to find the element within its children with the given condition
            mainWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);
            while (mainWnd == null)
            {
                mainWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);
            }
        }

        private static void Intro()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Thread.Sleep(sleepCount);

            write("Hello " + victim + ",");
            write("Would you like to play a game?");
            write("OK! Lets get started!");
        }

        private static void write(string s)
        {
            foreach (char c in s.ToCharArray())
            {
                write(c);
            }
            Console.WriteLine("");
            Thread.Sleep(longSleepCount);
        }

        private static void write(char c)
        {
            Console.Write(c);
            Thread.Sleep(sleepCount);
            if (c != 32)
            {
                Console.Beep();
            }
        }
    }
}