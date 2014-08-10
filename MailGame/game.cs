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
        private static string email = "rk598@saclink.csus.edu";
        private static string to = "My Boss";
        private static string subject = "How I feel";
        private static string message = "I Quit this job. You all can kiss my ass!";
        private static int sleepCount = 100;
        private static int longSleepCount = 3000;

        private static void Main(string[] args)
        {
            //Intro();
            play();
            Outro();
        }

        private static void Outro()
        {
            write("Wasn't that fun!");
            write("Would you like to play again?");
            Console.ReadKey();
        }

        private static void play()
        {
            AutomationElement mainWnd = null;
            AutomationElement emailWnd = null;
            getOutlookWindow(ref mainWnd);
            openNewEmail(ref emailWnd, mainWnd);
            setToField(emailWnd, to);
            // setSubjectField(emailWnd, subject);
            // setMessageField(emailWnd, message);
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            System.Windows.Forms.SendKeys.SendWait(subject);
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            System.Windows.Forms.SendKeys.SendWait(message);
            System.Windows.Forms.SendKeys.SendWait("^{s}");
            Thread.Sleep(2000);
            System.Windows.Forms.SendKeys.SendWait("%{F4}");
        }

        private static void setMessageField(AutomationElement emailWnd, string message)
        {
            PropertyCondition idCond = new PropertyCondition(AutomationElement.NameProperty, "Untitled Message");
            AutomationElement element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            while (element == null)
            {
                element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            }
            InsertTextUsingUIAutomation(element, message);
        }

        private static void setSubjectField(AutomationElement emailWnd, string subject)
        {
            PropertyCondition idCond = new PropertyCondition(AutomationElement.NameProperty, "Subject");
            AutomationElement element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            while (element == null)
            {
                element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            }

            InsertTextUsingUIAutomation(element, subject);
        }

        private static void setToField(AutomationElement emailWnd, string to)
        {
            PropertyCondition idCond = new PropertyCondition(AutomationElement.NameProperty, "To");
            AutomationElement element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            while (element == null)
            {
                element = emailWnd.FindFirst(TreeScope.Descendants, idCond);
            }
            if (element != null)
            {
                ValuePattern valPattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                valPattern.SetValue(to);
            }
        }

        private static void openNewEmail(ref AutomationElement emailWnd, AutomationElement mainWnd)
        {
            PropertyCondition idCond;
            AutomationElement element;

            idCond = new PropertyCondition(AutomationElement.NameProperty, "New Email");
            element = mainWnd.FindFirst(TreeScope.Descendants, idCond);
            while (element == null)
            {
                element = mainWnd.FindFirst(TreeScope.Descendants, idCond);
            }

            //Get invoke pattern and invoke it to press the button
            InvokePattern invPattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invPattern.Invoke();

            // reference the newly opened window
            //Create a property condition with the element's type
            PropertyCondition typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            //Create a property condition with the element's name
            PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "Untitled - Message (HTML) ");
            //Create the conjunction condition
            AndCondition andCondition = new AndCondition(typeCondition, nameCondition);
            //Ask the Desktop to find the element within its children with the given condition
            emailWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);
            while (emailWnd == null)
            {
                emailWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);
            }
        }

        private static void getOutlookWindow(ref AutomationElement mainWnd)
        {
            //Create a property condition with the element's type
            PropertyCondition typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            //Create a property condition with the element's name
            PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "Inbox - " + email + " - Outlook");
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

        private static void InsertTextUsingUIAutomation(AutomationElement element,
                                    string value)
        {
            try
            {
                // Validate arguments / initial setup
                if (value == null)
                    throw new ArgumentNullException(
                        "String parameter must not be null.");

                if (element == null)
                    throw new ArgumentNullException(
                        "AutomationElement parameter must not be null");

                // A series of basic checks prior to attempting an insertion.
                //
                // Check #1: Is control enabled?
                // An alternative to testing for static or read-only controls
                // is to filter using
                // PropertyCondition(AutomationElement.IsEnabledProperty, true)
                // and exclude all read-only text controls from the collection.
                if (!element.Current.IsEnabled)
                {
                    throw new InvalidOperationException(
                        "The control with an AutomationID of "
                        + element.Current.AutomationId.ToString()
                        + " is not enabled.\n\n");
                }

                // Check #2: Are there styles that prohibit us
                //           from sending text to this control?
                /* if (!element.Current.IsKeyboardFocusable)
                 {
                     throw new InvalidOperationException(
                         "The control with an AutomationID of "
                         + element.Current.AutomationId.ToString()
                         + "is read-only.\n\n");
                 }*/

                // Once you have an instance of an AutomationElement,
                // check if it supports the ValuePattern pattern.
                object valuePattern = null;

                // Control does not support the ValuePattern pattern
                // so use keyboard input to insert content.
                //
                // NOTE: Elements that support TextPattern
                //       do not support ValuePattern and TextPattern
                //       does not support setting the text of
                //       multi-line edit or document controls.
                //       For this reason, text input must be simulated
                //       using one of the following methods.
                //
                if (!element.TryGetCurrentPattern(
                    ValuePattern.Pattern, out valuePattern))
                {
                    // Set focus for input functionality and begin.
                    //element.SetFocus();

                    // Pause before sending keyboard input.
                    Thread.Sleep(100);

                    // Delete existing content in the control and insert new content.
                    System.Windows.Forms.SendKeys.SendWait(value);
                }
                // Control supports the ValuePattern pattern so we can
                // use the SetValue method to insert content.
                else
                {
                    // Set focus for input functionality and begin.
                    //element.SetFocus();

                    ((ValuePattern)valuePattern).SetValue(value);
                }
            }
            catch (ArgumentNullException ex)
            {
                Console.WriteLine(ex);
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}