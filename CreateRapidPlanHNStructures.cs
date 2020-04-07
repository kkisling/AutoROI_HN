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
    public class FindStructure
    {

    }


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
        const string BRAINSTEM_ID = "Brainstem";
        const string CORD_ID = "Spinal Cord";

        // Structure ID that will be written to plan
        const string PTVHIGH_ID = "PTV_HIGH";
        const string PTVINTDVH_ID = "PTV_INT_DVH";
        const string PTVINTOPT_ID = "PTV_INT_OPT";
        const string PTVLOWDVH_ID = "PTV_LOW_DVH";
        const string PTVLOWOPT_ID = "PTV_LOW_OPT";
        const string PAROTID_L_OPT_ID = "Parotid L OAR";
        const string PAROTID_R_OPT_ID = "Parotid R OAR";
        const string PHARYNX_OPT_ID = "OAR_Pharynx";
        const string RING3_10_ID = "3_10mm RING";
        const string BRAINSTEM3mm_ID = "Brainstem 3PRV";
        const string CORD5mm_ID = "Spinal Cord 5PRV";

        // Structure ID for temporary structures, will not be written to plan
        const string BODY3mm_ID = "Body_3mm";
        const string PTVINT_ID = "PTV_INT";
        const string PTVLOW_ID = "PTV_LOW";
        const string PTVHIGHINT_ID = "PTV_HIGH+INT";
        const string PTVALL_ID = "PTV_ALL";
        const string PTVHIGH3mm_ID = "PTV_HIGH+3mm";
        const string PTVHIGHINT3mm_ID = "PTV_HIGH+INT+3mm";
        const string PTVALL3mm_ID = "PTV_ALL+3mm";
        const string PTVALL10mm_ID = "PTV_ALL+10mm";
        const string RING3_10PRE_ID = "3_10mm RING_pre";


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

                List<string> misingrois = new List<string>(){ };

                Structure body = ss.Structures.FirstOrDefault(x => x.Id == BODY_ID);
                misingrois = CheckBaseROI(body, BODY_ID, SCRIPT_NAME, misingrois);

                Structure ptv_highclin = ss.Structures.FirstOrDefault(x => x.Id == PTVHIGHclin_ID);
                misingrois = CheckBaseROI(ptv_highclin, PTVHIGHclin_ID, SCRIPT_NAME, misingrois);

                Structure ptv_intclin = ss.Structures.FirstOrDefault(x => x.Id == PTVINTclin_ID);
                misingrois = CheckBaseROI(ptv_intclin, PTVINTclin_ID, SCRIPT_NAME, misingrois);

                Structure ptv_lowclin = ss.Structures.FirstOrDefault(x => x.Id == PTVLOWclin_ID);
                misingrois = CheckBaseROI(ptv_lowclin, PTVLOWclin_ID, SCRIPT_NAME, misingrois);

                if (misingrois.Count > 0)
                {
                    foreach (string element in misingrois)
                    {
                        MessageBox.Show(string.Format("'{0}' not found! Exiting script", element), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                    return;
                }


                //============================
                // CHECK if PTV_HIGH already exists. If it does exist, show warning and exit code
                //============================
                // Do this because other structures that this code makes rely on it

                Structure check_PTVHigh = ss.Structures.FirstOrDefault(x => x.Id == PTVHIGH_ID);
                if (check_PTVHigh != null)
                {
                    MessageBox.Show(string.Format("'{0}' already exists! Please delete or rename '{0}' and re-run script.", PTVHIGH_ID,
                        PTVHIGH_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }


                //============================
                // FIND OARs. If they don't exist, show warning and skip. CHECK for OAR_planning. If it exists, show warning and skip
                //============================

                // find Parotid L
                Structure parotid_L = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_L_ID);
                bool flagOAR_parotidL = CheckOAR(PAROTID_L_ID, parotid_L, SCRIPT_NAME, ss, PAROTID_L_OPT_ID);

                // find Parotid R
                Structure parotid_R = ss.Structures.FirstOrDefault(x => x.Id == PAROTID_R_ID);
                bool flagOAR_parotidR = CheckOAR(PAROTID_R_ID, parotid_R, SCRIPT_NAME, ss, PAROTID_R_OPT_ID);

                // find Pharynx
                Structure pharynx = ss.Structures.FirstOrDefault(x => x.Id == PHARYNX_ID);
                bool flagOAR_pharynx = CheckOAR(PHARYNX_ID, pharynx, SCRIPT_NAME, ss, PHARYNX_OPT_ID);

                // find Brainstem
                Structure brainstem = ss.Structures.FirstOrDefault(x => x.Id == BRAINSTEM_ID);
                bool flagOAR_brainstem = CheckOAR(BRAINSTEM_ID, brainstem, SCRIPT_NAME, ss, BRAINSTEM3mm_ID);

                // find Spinal Cord
                Structure cord = ss.Structures.FirstOrDefault(x => x.Id == CORD_ID);
                bool flagOAR_cord = CheckOAR(CORD_ID, cord, SCRIPT_NAME, ss, CORD5mm_ID);


                //============================
                // CHECK if structures to be written already exist. If they do exist, show warning and skip creating those
                //============================

                // Check for PTV_INT_DVH
                bool flagPTV_INTDVH = CheckNewROI(PTVINTDVH_ID, ss, SCRIPT_NAME);

                // Check for PTV_INT_OPT
                bool flagPTV_INTOPT = CheckNewROI(PTVINTOPT_ID, ss, SCRIPT_NAME);

                // Check for PTV_LOW_DVH
                bool flagPTV_LOWDVH = CheckNewROI(PTVLOWDVH_ID, ss, SCRIPT_NAME);

                // Check for PTV_LOW_OPT
                bool flagPTV_LOWOPT = CheckNewROI(PTVLOWOPT_ID, ss, SCRIPT_NAME);

                // Check for 3_10mm RING
                bool flagPTV_RING = CheckNewROI(RING3_10_ID, ss, SCRIPT_NAME);


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
                    // 8 GENERATE 3_10mm RING, which is PTV_ALL+10mm Subtracting PTV_ALL+3mm
                    //============================       
                    if (flagPTV_RING == true)
                    {
                        Structure ptv_all3mm = ss.AddStructure("CONTROL", PTVALL3mm_ID);
                        Structure ptv_all10mm = ss.AddStructure("CONTROL", PTVALL10mm_ID);
                        Structure ring_3_10pre = ss.AddStructure("CONTROL", RING3_10PRE_ID);
                        Structure ring_3_10 = ss.AddStructure("CONTROL", RING3_10_ID);
                        ptv_all3mm.SegmentVolume = ptv_all.Margin(3.0);
                        ptv_all10mm.SegmentVolume = ptv_all.Margin(10.0);
                        ring_3_10pre.SegmentVolume = ptv_all10mm.Sub(ptv_all3mm);
                        ring_3_10.SegmentVolume = ring_3_10pre.And(body_3mm);

                        string message8 = string.Format("{0} was created with volume {1}.", ring_3_10.Id, ring_3_10.Volume);

                        MessageBox.Show(message8);

                        ss.RemoveStructure(ptv_all3mm);
                        ss.RemoveStructure(ptv_all10mm);
                        ss.RemoveStructure(ring_3_10pre);
                    }


                    //============================
                    // 101 SUBTRACT from the parotids (L and R) the PTVs
                    //============================
                    if (flagOAR_parotidL == true)
                    {
                        Structure parotid_L_oar = ss.AddStructure("AVOIDANCE", PAROTID_L_OPT_ID);
                        parotid_L_oar.SegmentVolume = parotid_L.Sub(ptv_all);

                        string message101_L = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                                parotid_L.Id, parotid_L_oar.Id, parotid_L.Volume, parotid_L_oar.Volume);
                        MessageBox.Show(message101_L);
                    }

                    if (flagOAR_parotidR == true)
                    {
                        Structure parotid_R_oar = ss.AddStructure("AVOIDANCE", PAROTID_R_OPT_ID);
                        parotid_R_oar.SegmentVolume = parotid_R.Sub(ptv_all);

                        string message101_R = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                                parotid_R.Id, parotid_R_oar.Id, parotid_R.Volume, parotid_R_oar.Volume);
                        MessageBox.Show(message101_R);
                    }

                    //============================
                    // 102 SUBTRACT from the pharynx the PTVs
                    //============================
                    if (flagOAR_pharynx == true)
                    {
                        Structure pharynx_oar = ss.AddStructure("AVOIDANCE", PHARYNX_OPT_ID);
                        // Check if a high res structure. If so, try copying to a new, non high res structure
                        if (pharynx.IsHighResolution == true)
                        {
                            pharynx_oar.ConvertToHighResolution();
                            // Can I skip specifying that?
                            const string PTV_ALL_HIGHRES_ID = "PTV_ALL_HighRes";
                            Structure ptv_all_highres = ss.AddStructure("CONTROL", PTV_ALL_HIGHRES_ID);
                            ptv_all_highres.SegmentVolume = ptv_all.SegmentVolume;
                            ptv_all_highres.ConvertToHighResolution();
                            pharynx_oar.SegmentVolume = pharynx.Sub(ptv_all_highres);
                            ss.RemoveStructure(ptv_all_highres);

                            string message102HR = string.Format("{0} was a high resolution structure.", pharynx.Id);
                            MessageBox.Show(message102HR);
                        }
                        else
                        {
                            pharynx_oar.SegmentVolume = pharynx.Sub(ptv_all);
                        }

                        string message102 = string.Format("{0} volume = {2}\n{1} created with volume = {3}",
                                pharynx.Id, pharynx_oar.Id, pharynx.Volume, pharynx_oar.Volume);
                        MessageBox.Show(message102);
                    }

                    //============================
                    // 103 EXPAND the Brainstem 3mm
                    //============================
                    if (flagOAR_brainstem == true)
                    {
                        Structure brainstem_PRV = ss.AddStructure("AVOIDANCE", BRAINSTEM3mm_ID);
                        brainstem_PRV.SegmentVolume = brainstem.Margin(3.0);

                        string message103 = string.Format("{0} volume = {2}\n{1} created with volume {3}",
                            brainstem.Id, brainstem_PRV.Id, brainstem.Volume, brainstem_PRV.Volume);
                        MessageBox.Show(message103);
                    }

                    //============================
                    // 104 EXPAND the Brainstem 3mm
                    //============================
                    if (flagOAR_cord == true)
                    {
                        Structure cord_PRV = ss.AddStructure("AVOIDANCE", CORD5mm_ID);
                        cord_PRV.SegmentVolume = cord.Margin(5.0);

                        string message104 = string.Format("{0} volume = {2}\n{1} created with volume {3}",
                            cord.Id, cord_PRV.Id, cord.Volume, cord_PRV.Volume);
                        MessageBox.Show(message104);
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

        // HELPER METHODS****************

        private static List<string> CheckBaseROI(Structure baseroi, string BASEROI_ID, string SCRIPT_NAME, List<string> missingrois)
        {
            if (baseroi == null)
            {
                missingrois.Add(BASEROI_ID);
            }
            return missingrois;
        }

        private static bool CheckOAR(string OAR_ID, Structure oar, string SCRIPT_NAME, StructureSet ss, string OAR_OPT_ID)
        {
            // Check for OAR needed to make optimization OAR
            bool flagOAR = true;
            if (oar == null)
            {
                flagOAR = false;
                MessageBox.Show(string.Format("'{0}' not found! '{1}' will not be created.", OAR_ID,
                    OAR_OPT_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                // Check if optimization OAR already exists
                Structure check_OAR = ss.Structures.FirstOrDefault(x => x.Id == OAR_OPT_ID);
                if (check_OAR != null)
                {
                    flagOAR = false;
                    MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be re-created.", OAR_OPT_ID,
                        OAR_OPT_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            return flagOAR;
        }

        private static bool CheckNewROI(string ROI_ID, StructureSet ss, string SCRIPT_NAME)
        {
            Structure check_ROI = ss.Structures.FirstOrDefault(x => x.Id == ROI_ID);
            bool flagROI = true;
            if (check_ROI != null)
            {
                flagROI = false;
                MessageBox.Show(string.Format("'{0}' already exists! '{0}' will not be re-created.", ROI_ID,
                    ROI_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            return flagROI;
        }

    }
}

