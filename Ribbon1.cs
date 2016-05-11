using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

// TODO:  suivez ces étapes pour activer l'élément (XML) Ruban :

// 1 : Copiez le bloc de code suivant dans la classe ThisAddin, ThisWorkbook ou ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Créez des méthodes de rappel dans la région "Rappels du ruban" de cette classe pour gérer les actions des utilisateurs,
//    comme les clics sur un bouton. Remarque : si vous avez exporté ce ruban à partir du Concepteur de ruban,
//    vous devrez déplacer votre code des gestionnaires d'événements vers les méthodes de rappel et modifiez le code pour qu'il fonctionne avec
//    le modèle de programmation d'extensibilité du ruban (RibbonX).

// 3. Assignez les attributs aux balises de contrôle dans le fichier XML du ruban pour identifier les méthodes de rappel appropriées dans votre code.  

// Pour plus d'informations, consultez la documentation XML du ruban dans l'aide de Visual Studio Tools pour Office.


namespace MeetingAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region Membres IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MeetingAddIn.Ribbon1.xml");
        }

        #endregion

        #region Rappels du ruban
        //Créez des méthodes de rappel ici. Pour plus d'informations sur l'ajout de méthodes de rappel, consultez http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Programmes d'assistance

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        public void OnCallCreateAppointment(Office.IRibbonControl control)
        {
            try
            {
                var outlookApp = new Outlook.Application();
                Outlook.AppointmentItem newAppointment = outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);
                newAppointment.Start = DateTime.Now.AddHours(2);
                newAppointment.End = DateTime.Now.AddHours(3);
                newAppointment.Location = "ConferenceRoom #2345";
                newAppointment.Body =
                    "We will discuss progress on the group project.";
                
                newAppointment.AllDayEvent = false;
                newAppointment.Subject = "Meet-o Project";
                newAppointment.Recipients.Add("pascal.verdage@sgcib.com");
                Outlook.Recipients sentTo = newAppointment.Recipients;
                Outlook.Recipient sentInvite = null;
                sentInvite = sentTo.Add("Holly Holt");
                sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                    .olRequired;
                sentInvite = sentTo.Add("David Junca ");
                sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                    .olOptional;
                sentTo.ResolveAll();
                newAppointment.Save();
                newAppointment.Display(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }
    }
}
