using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using MSXML;
using XmlService;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;

namespace BroadcastExtension
{

    [ComVisible(true)]
    [ProgId("BroadcastExtension.BroadcastExtension")]
    public class BroadcastExtension : IWorkflowExtension
    {

        private INautilusServiceProvider sp;
        public void Execute(ref LSExtensionParameters Parameters)
        {

            string tableName = Parameters["TABLE_NAME"];
            sp = Parameters["SERVICE_PROVIDER"];
            var Id = Parameters["WORKFLOW_NODE_ID"];
            var rs = Parameters["RECORDS"];


            var bbb = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(bbb);
            var xmlProcessor = Utils.GetXmlProcessor(sp);
            var dal = new DataLayer();
            dal.Connect();

            WorkflowNode workflowNode = dal.GetWorkFlowNodeByID(Id);
            string extensionname = workflowNode.LONG_NAME.Replace("Extension - ", "");
            Extension extension = dal.GetExtensionByName(extensionname);
            string eventsToFire = extension.Description;
            var eventToFireSplit = new List<string>(eventsToFire.Split(';'));
            var entityName = rs.Fields["NAME"].Value;


            switch (tableName)
            {
                case "SDG":
                    Sdg sdg = dal.GetSdgByName(entityName);
                    CaseSdg(sdg, xmlProcessor, eventToFireSplit);
                    break;
                case "SAMPLE":
                    Sample sample = dal.GetSampleByName(entityName);
                    CaseSample(sample, xmlProcessor, eventToFireSplit);
                    break;
                case "ALIQUOT":
                    Aliquot aliquot = dal.GetAliquotByName(entityName);
                    CaseAliquot(aliquot, xmlProcessor, eventToFireSplit);
                    break;
                case "TEST":
                    Test test = dal.GetTestByName(entityName);
                    CaseTest(test, xmlProcessor, eventToFireSplit);
                    break;
                case "RESULT":
                    Result result = dal.GetResultByName(entityName);
                    CaseResult(result, xmlProcessor, eventToFireSplit);
                    break;
            }
        }

        private void EntityFireEvent(INautilusProcessXML xmlProcessror, string tableName, long EntityID, string EventName)
        {
            var res = new DOMDocument();
            var fireEventXml = new FireEventXmlHandler(sp);
            fireEventXml.CreateFireEventXml(tableName, EntityID, EventName);
            fireEventXml.ProcssXml();

            //doc.save(@"C:\Users\hilae\Desktop\xmlPrint\" + tableName + EntityID + "doc.xmlProcessror");
            //res.save(@"C:\Users\hilae\Desktop\xmlPrint\" + tableName + EntityID + "res.xmlProcessror");
        }

        private void CaseResult(Result result, INautilusProcessXML xml, IEnumerable<string> eventToFireSplit)
        {
            if (result.Status == "X") return;

            foreach (var eventName in eventToFireSplit)
            {
                if (result.EVENTS != null && result.EVENTS.Contains(eventName + ","))
                {
                    EntityFireEvent(xml, "RESULT", result.ResultId, eventName);
                }
            }

        }

        private void CaseTest(Test test, INautilusProcessXML xml, IEnumerable<string> eventToFireSplit)
        {
            if (test.STATUS == "X") return;

            foreach (var eventName in eventToFireSplit)
            {
                if (test.EVENTS != null && test.EVENTS.Contains(eventName + ","))
                {
                    EntityFireEvent(xml, "TEST", test.TEST_ID, eventName);
                }
                var results = test.Results;
                foreach (var result in results)
                {
                    CaseResult(result, xml, eventToFireSplit);
                }
            }

        }

        private void CaseAliquot(Aliquot aliquot, INautilusProcessXML xml, IEnumerable<string> eventToFireSplit)
        {
            if (aliquot.Status == "X") return;

            foreach (var EventName in eventToFireSplit)
            {
                if (aliquot.EVENTS != null && aliquot.EVENTS.Contains(EventName + ","))
                {
                    EntityFireEvent(xml, "ALIQUOT", aliquot.AliquotId, EventName);
                }
                var tests = aliquot.Tests;
                foreach (var test in tests)
                {
                    CaseTest(test, xml, eventToFireSplit);
                }
                foreach (Aliquot child in aliquot.Children)
                {
                    CaseAliquot(child, xml, eventToFireSplit);
                }
            }

        }

        private void CaseSample(Sample sample, INautilusProcessXML xml, IEnumerable<string> eventToFireSplit)
        {

            foreach (var EventName in eventToFireSplit)
            {
                if (sample.EVENTS != null && sample.EVENTS.Contains(EventName + ","))
                {
                    EntityFireEvent(xml, "SAMPLE", sample.SampleId, EventName);
                }
                var aliquots = sample.Aliqouts;

                foreach (var aliquot in aliquots)
                {
                    if (aliquot.Parent.Count == 0)
                    {
                        CaseAliquot(aliquot, xml, eventToFireSplit);
                    }

                }
            }

        }

        private void CaseSdg(Sdg sdg, INautilusProcessXML xml, IEnumerable<string> eventToFireSplit)
        {
            foreach (var EventName in eventToFireSplit)
            {
                if (sdg.EVENTS != null && sdg.EVENTS.Contains(EventName + ","))
                {
                    EntityFireEvent(xml, "SDG", sdg.SdgId, EventName);
                }
                var samples = sdg.Samples;
                foreach (var sample in samples)
                {
                    CaseSample(sample, xml, eventToFireSplit);
                }
            }

        }


    }
}
