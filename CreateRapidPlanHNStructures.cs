////////////////////////////////////////////////////////////////////////////////
// CreateRapidPlanHNStructures.cs
// Edited from: CreateOptStructures.cs
//
//
// Applies to:
//      Eclipse Scripting API
//          15.1.1
//          15.5
//
////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

//[assembly: AssemblyVersion("1.0.0.1")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        const string SCRIPT_NAME = "HN RapidPlan Structures Script";
        // Structure ID that should already exist in plan
        const string BODY_ID = "Body";
        const string PTVHIGHclin_ID = "PTV_7000";
        const string PTVINTclin_ID = "PTV_6125";
        const string PTVLOWclin_ID = "PTV_5600";
        const string PAROTID_L_ID = "Parotid L";
        const string PAROTID_R_ID = "Parotid R";
        const string PHARYNX_ID = "Pharynx";
        // Structure ID that will be written to plan
        const string PTVHIGH_ID = "PTV_HIGH";
        const string PTVINTDVH_ID = "PTV_INT_DVH";
        const string PTVINTOPT_ID = "PTV_INT_OPT";
        const string PTVLOWDVH_ID = "PTV_LOW_DVH";
        const string PTVLOWOPT_ID = "PTV_LOW_OPT";
        const string PAROTID_L_OPT_ID = "Parotid L OAR";
        const string PAROTID_R_OPT_ID = "Parotid R OAR";
        const string PHARYNX_OPT_ID = "OAR_Pharynx";
        // Structure ID for temporary structures, will not be written to plan
        const string BODY3mm_ID = "Body_3mm";
        const string PTVINT_ID = "PTV_INT";
        const string PTVLOW_ID = "PTV_LOW";
        const string PTVHIGHINT_ID = "PTV_HIGH+INT";
        const string PTVALL_ID = "PTV_all";
        const string PTVHIGH3mm_ID = "PTV_HIGH+3mm";
        const string PTVHIGHINT3mm_ID = "PTV_HIGH+INT+3mm";

        public Script()
        {
        }

        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {

            //============================
            // FIND Structure Set. If patient or structure set don't exist, show warning and exit code
            //============================

            if (context.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            StructureSet ss = context.StructureSet;

            //============================
            // FIND Body and PTVs. If they don't exist, show warning and exit code
            //============================

            Structure body = ss.Structures.FirstOrDefault(x => x.Id == BODY_ID);
            if (body == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", BODY_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            Structure ptv_highclin = ss.Structures.FirstOrDefault(x => x.Id == PTVHIGHclin_ID);
            if (ptv_highclin == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", PTVHIGHclin_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            Structure ptv_intclin = ss.Structures.FirstOrDefault(x => x.Id == PTVINTclin_ID);
            if (ptv_intclin == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", PTVINTclin_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            Structure ptv_lowclin = ss.Structures.FirstOrDefault(x => x.Id == PTVLOWclin_ID);
            if (ptv_lowclin == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", PTVLOWclin_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            //============================
            // FIND OARs. If they don't exist, show warning and skip these
            //============================

            bool flagOAR_parotidL = true;
            bool flagOAR_parotidR = true;
            bool flagOAR_pharynx = true;

            // find Parotid L
            Structure parotid_L = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_L_ID);
            if (parotid_L == null)
            {
                flagOAR_parotidL = false;
                MessageBox.Show(string.Format("'{0}' not found! OK to continue? '{1}' will not be created.", PAROTID_L_ID,
                    PAROTID_L_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // find Parotid R
            Structure parotid_R = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_R_ID);
            if (parotid_R == null)
            {
                flagOAR_parotidR = false;
                MessageBox.Show(string.Format("'{0}' not found! OK to continue? '{1}' will not be created.", PAROTID_R_ID,
                    PAROTID_R_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // find Pharynx
            Structure pharynx = ss.Structures.FirstOrDefault(x => x.Id == PHARYNX_ID);
            if (pharynx == null)
            {
                flagOAR_pharynx = false;
                MessageBox.Show(string.Format("'{0}' not found! OK to continue? '{1}' will not be created.", PHARYNX_ID,
                    PHARYNX_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
                return;
            }

            //============================
            // CHECK if PTV_HIGH already exist. If it does exist, show warning and exit code
            //============================

            Structure check_PTVHigh = ss.Structures.FirstOrDefault(x => x.Id == PTVHIGH_ID);
            if (check_PTVHigh != null)
            {
                MessageBox.Show(string.Format("'{0}' already exists! Please delete or rename '{0}' and re-run script.", PTVHIGH_ID,
                    PTVHIGH_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            //============================
            // CHECK if structures to be written already exist. If they do exist, show warning and skip creating those
            //============================

            bool flagPTV_INTDVH = true;
            bool flagPTV_INTOPT = true;
            bool flagPTV_LOWDVH = true;
            bool flagPTV_LOWOPT = true;

            // Check for PTV_INT_DVH
            Structure check_PTVIntDVH = ss.Structures.FirstOrDefault(x => x.Id == PTVINTDVH_ID);
            if (check_PTVIntDVH != null)
            {
                flagPTV_INTDVH = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PTVINTDVH_ID,
                    PTVINTDVH_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // Check for PTV_INT_OPT
            Structure check_PTVIntOPT = ss.Structures.FirstOrDefault(x => x.Id == PTVINTOPT_ID);
            if (check_PTVIntOPT != null)
            {
                flagPTV_INTOPT = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PTVINTOPT_ID,
                    PTVINTOPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // Check for PTV_LOW_DVH
            Structure check_PTVLowDVH = ss.Structures.FirstOrDefault(x => x.Id == PTVLOWDVH_ID);
            if (check_PTVLowDVH != null)
            {
                flagPTV_LOWDVH = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PTVLOWDVH_ID,
                    PTVLOWDVH_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // Check for PTV_LOW_OPT
            Structure check_PTVLowOPT = ss.Structures.FirstOrDefault(x => x.Id == PTVLOWOPT_ID);
            if (check_PTVLowOPT != null)
            {
                flagPTV_LOWOPT = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PTVLOWOPT_ID,
                    PTVLOWOPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // Check for Parotid L OAR
            Structure check_ParotidLOAR = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_L_OPT_ID);
            if (check_ParotidLOAR != null)
            {
                flagOAR_parotidL = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PAROTID_L_OPT_ID,
                    PAROTID_L_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }

            // Check for Parotid R OAR
            Structure check_ParotidROAR = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_R_OPT_ID);
            if (check_ParotidROAR != null)
            {
                flagOAR_parotidR = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PAROTID_R_OPT_ID,
                    PAROTID_R_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }


            // Check for Pharynx
            Structure check_Pharynx = ss.Structures.FirstOrDefault(x => x.Id == PHARYNX_OPT_ID);
            if (check_Pharynx != null)
            {
                flagOAR_pharynx = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be created. OK to continue? ", PHARYNX_OPT_ID,
                    PHARYNX_OPT_ID), SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            }
            

            context.Patient.BeginModifications();   // enable writing with this script.

            try
            {
                //============================
                // 1 GENERATE 3mm contraction of body
                //============================
                Structure body_3mm = ss.AddStructure("CONTROL", BODY3mm_ID);
                body_3mm.SegmentVolume = body.Margin(-3.0);

                string message1 = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                        body.Id, body_3mm.Id, body.Volume, body_3mm.Volume);
                MessageBox.Show(message1);

                //============================
                // 2 GENERATE PTV_HIGH, PTV_INT, and PTV_LOW, from which is subtracted 3mm from body
                //============================       

                Structure ptv_high = ss.AddStructure("PTV", PTVHIGH_ID);
                Structure ptv_int = ss.AddStructure("PTV", PTVINT_ID);
                Structure ptv_low = ss.AddStructure("PTV", PTVLOW_ID);
                ptv_high.SegmentVolume = ptv_highclin.And(body_3mm);
                ptv_int.SegmentVolume = ptv_intclin.And(body_3mm);
                ptv_low.SegmentVolume = ptv_lowclin.And(body_3mm);

                string message2 = string.Format("{0} volume = {6}\n{1} created with volume = {7}\n{2} volume = " +
                    "{8}\n{3} created with volume = {9}\n{4} volume = {10}\n{5} created with volume {11}",
                        ptv_highclin.Id, ptv_high.Id, ptv_intclin.Id, ptv_int.Id, ptv_lowclin.Id, ptv_low.Id,
                        ptv_highclin.Volume, ptv_high.Volume, ptv_intclin.Volume, ptv_int.Volume, ptv_lowclin.Volume, ptv_low.Volume);
                MessageBox.Show(message2);

                //============================
                // 3 GENERATE PTV_ALL, which is PTV_HIGH + PTV_INT + PTV+LOW
                //============================       
                Structure ptv_highint = ss.AddStructure("PTV", PTVHIGHINT_ID);
                Structure ptv_all = ss.AddStructure("PTV", PTVALL_ID);
                ptv_highint.SegmentVolume = ptv_high.Or(ptv_int);
                ptv_all.SegmentVolume = ptv_highint.Or(ptv_low);

                string message3 = string.Format("{0} volume = {5}\n{1} volume = {6}\n{2} created with volume = {7}" +
                    "\n{3} volume = {8}\n{4} created with volume = {9}",
                        ptv_high.Id, ptv_int.Id, ptv_highint.Id, ptv_low.Id, ptv_all.Id,
                        ptv_high.Volume, ptv_int.Volume, ptv_highint.Volume, ptv_low.Volume, ptv_all.Volume);
                MessageBox.Show(message3);

                //============================
                // 4 GENERATE PTV_INT_DVH, which is PTV_INT Subtracting PTV_HIGH
                //============================      
                if (flagPTV_INTDVH == true)
                {
                    Structure ptv_intdvh = ss.AddStructure("PTV", PTVINTDVH_ID);
                    ptv_intdvh.SegmentVolume = ptv_int.Sub(ptv_high);

                    string message4 = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                            ptv_int.Id, ptv_intdvh.Id, ptv_int.Volume, ptv_intdvh.Volume);
                    MessageBox.Show(message4);
                }
                
                //============================
                // 5 GENERATE PTV_INT_OPT, which is PTV_INT Subtracting PTV_HIGH+3mm
                //============================  
                if (flagPTV_INTOPT == true)
                {
                    Structure ptv_high3mm = ss.AddStructure("PTV", PTVHIGH3mm_ID);
                    Structure ptv_intopt = ss.AddStructure("PTV", PTVINTOPT_ID);
                    ptv_high3mm.SegmentVolume = ptv_high.Margin(3.0);
                    ptv_intopt.SegmentVolume = ptv_int.Sub(ptv_high3mm);

                    string message5 = string.Format("{0} volume = {4}\n{1} created with volume = {5}\n" +
                        "{2} volume = {6}\n{3} created with volume = {7}",
                            ptv_high.Id, ptv_high3mm.Id, ptv_int.Id, ptv_intopt.Id,
                            ptv_high.Volume, ptv_high3mm.Volume, ptv_int.Volume, ptv_intopt.Volume);
                    MessageBox.Show(message5);

                    ss.RemoveStructure(ptv_high3mm);
                }

                //============================
                // 6 GENERATE PTV_LOW_DVH, which is PTV_LOW Subtracting PTV_HIGH+INT
                //============================   
                if (flagPTV_LOWDVH == true)
                {
                    Structure ptv_lowdvh = ss.AddStructure("PTV", PTVLOWDVH_ID);
                    ptv_lowdvh.SegmentVolume = ptv_low.Sub(ptv_highint);

                    string message6 = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                            ptv_low.Id, ptv_lowdvh.Id, ptv_low.Volume, ptv_lowdvh.Volume);
                    MessageBox.Show(message6);
                }

                //============================
                // 7 GENERATE PTV_LOW_OPT, which is PTV_LOW Subtracting PTV_HIGH+INT+3mm
                //============================       
                if (flagPTV_LOWOPT == true)
                {
                    Structure ptv_highint3mm = ss.AddStructure("PTV", PTVHIGHINT3mm_ID);
                    Structure ptv_lowopt = ss.AddStructure("PTV", PTVLOWOPT_ID);
                    ptv_highint3mm.SegmentVolume = ptv_highint.Margin(3.0);
                    ptv_lowopt.SegmentVolume = ptv_low.Sub(ptv_highint3mm);

                    string message7 = string.Format("{0} volume = {4}\n{1} created with volume = {5}\n" +
                        "{2} volume = {6}\n{3} created with volume = {7}",
                            ptv_highint.Id, ptv_highint3mm.Id, ptv_low.Id, ptv_lowopt.Id,
                            ptv_highint.Volume, ptv_highint3mm.Volume, ptv_low.Volume, ptv_lowopt.Volume);
                    MessageBox.Show(message7);

                    ss.RemoveStructure(ptv_highint3mm);
                }

                //============================
                // 8 SUBTRACT from the parotids (L and R) the PTVs
                //============================
                if (flagOAR_parotidL == true)
                {
                    Structure parotid_L_oar = ss.AddStructure("AVOIDANCE", PAROTID_L_OPT_ID);
                    parotid_L_oar.SegmentVolume = parotid_L.Sub(ptv_all);

                    string message8_L = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                            parotid_L.Id, parotid_L_oar.Id, parotid_L.Volume, parotid_L_oar.Volume);
                    MessageBox.Show(message8_L);
                }

                if (flagOAR_parotidR == true)
                {
                    Structure parotid_R_oar = ss.AddStructure("AVOIDANCE", PAROTID_R_OPT_ID);
                    parotid_R_oar.SegmentVolume = parotid_R.Sub(ptv_all);

                    string message8_R = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                            parotid_R.Id, parotid_R_oar.Id, parotid_R.Volume, parotid_R_oar.Volume);
                    MessageBox.Show(message8_R);
                }

                //============================
                // 9 SUBTRACT from the pharynx the PTVs
                //============================

                if (flagOAR_pharynx == true)
                {
                    Structure pharynx_oar = ss.AddStructure("AVOIDANCE", PHARYNX_OPT_ID);
                    pharynx_oar.SegmentVolume = pharynx.Sub(ptv_all); 

                    string message9 = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                            pharynx.Id, pharynx_oar.Id, pharynx.Volume, pharynx_oar.Volume);
                    MessageBox.Show(message9);
                }

                //============================
                // PRINT out all relevant structures. Delete the rest! **** FOR FINAL VERSION
                //============================

                ss.RemoveStructure(body_3mm);
                ss.RemoveStructure(ptv_int);
                ss.RemoveStructure(ptv_low);
                ss.RemoveStructure(ptv_highint);
                ss.RemoveStructure(ptv_all);


            }
            catch (Exception e)
            {
                string messageError = string.Format("{0} Exception caught.", e);
                MessageBox.Show(messageError);
            }
        }
    }
}

