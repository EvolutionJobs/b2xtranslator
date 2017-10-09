using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(179)]
    public class AccentCallout90Type : ShapeType
    {
        public AccentCallout90Type()
        {
            this.ShapeConcentricFill = true;
            
            this.Joins = JoinStyle.miter;        

            this.Path = "m@0@1l@2@3nfem,l21600,r,21600l,21600nsxe";

            this.Formulas = new List<string>();                     
            this.Formulas.Add("val #0");                       
            this.Formulas.Add("val #1");                       
            this.Formulas.Add("val #2");                       
            this.Formulas.Add("val #3");
            this.AdjustmentValues = "-1800,24300,-1800,4050"; 
            this.ConnectorLocations = "@0,@1;10800,0;10800,21600;0,10800;21600,10800";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,#1";
            this.Handles.Add(HandleOne);

            Handle HandleTwo = new Handle();
            HandleTwo.position="#2,#3";
            this.Handles.Add(HandleTwo);
        }
    }
}
