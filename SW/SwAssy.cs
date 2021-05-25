﻿using System;
using SldWorks;

namespace SW
{
    public class SwAssy
    {
        public string Name;
        public string Description;
        public string CompanyNo;
        private SldWorks.ModelDoc2 swModel;
       
        private SldWorks.Configuration swConf;
        public string ConfName;
        public SldWorks.Component2 Comp;
        public SwAssy(SldWorks.Component2 comp)
        {
            if (comp != null)
            {
                Comp = comp;
                swModel = ( SldWorks.ModelDoc2) comp.GetModelDoc2();
                
                Console.WriteLine("New assembly" + swModel.CustomInfo2[swConf.Name, "Description"]);
                Name = swModel.GetTitle();
                swConf = (SldWorks.Configuration) swModel.GetActiveConfiguration();
                ConfName = swConf.Name;
                Description = swModel.CustomInfo2[swConf.Name, "Description"];
                CompanyNo = swModel.CustomInfo2[swConf.Name, "Company No"];
            }
        }
        
        public string GetProperty(string key)
        {
            return ((ModelDoc2)Comp.GetModelDoc2()).CustomInfo2[ConfName, key];
        }
    }
}