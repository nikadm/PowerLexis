using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddIn
{
    public partial class Ribbon1
    {
        private Presentation newPreset;

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            newPreset = CreateNewPPT();
        }

        public Presentation CreateNewPPT()
        {

            var appPPT = new Application();
            var pptPresent = appPPT.Presentations.Add();

            var pptLayout = pptPresent.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutChartAndText];
            var pptLayout1 = pptPresent.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
            var pptLayout2 = pptPresent.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTable];

            pptPresent.Slides.AddSlide((pptPresent.Slides.Count + 1), pptLayout);
            
            pptPresent.Slides.AddSlide((pptPresent.Slides.Count + 1), pptLayout1);
            pptPresent.Slides.AddSlide((pptPresent.Slides.Count + 1), pptLayout2);

            return pptPresent;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var pptLayout = newPreset.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
            newPreset.Slides.AddSlide((newPreset.Slides.Count + 1), pptLayout);
        }

       
    }
}
