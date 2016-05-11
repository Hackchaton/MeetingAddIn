using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MeetingAddIn
{
    partial class FormRegion1
    {
        #region Fabrique de zones de formulaire

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("MeetingAddIn.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Se produit avant l'initialisation de la zone de formulaire.
            // Pour empêcher l'affichage de la zone de formulaire, définissez e.Cancel à true.
            // Utilisez e.OutlookItem pour obtenir une référence à l'élément Outlook actuel.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
               
            }
        }

        #endregion

        // Se produit avant l'affichage de la zone de formulaire.
        // Utilisez this.OutlookItem pour obtenir une référence à l'élément Outlook actuel.
        // Utilisez this.OutlookFormRegion pour obtenir une référence à la zone de formulaire.
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
        }

        // Se produit à la fermeture de la zone de formulaire.
        // Utilisez this.OutlookItem pour obtenir une référence à l'élément Outlook actuel.
        // Utilisez this.OutlookFormRegion pour obtenir une référence à la zone de formulaire.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
