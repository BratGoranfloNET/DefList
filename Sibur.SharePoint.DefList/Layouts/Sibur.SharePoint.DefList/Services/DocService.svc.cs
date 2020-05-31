using System.ServiceModel.Activation;

namespace Sibur.SharePoint.DefList
{
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
    public class DocService : IDocService
    {
        public string DocServiceCall(string SampleValue)
        {
            return "ok";
        }

        public ResponseModel GenerateExcelFileGET(string listDefUrlOrID, string listDefItemId, string listOrgUrlOrID, string listOrgItemId)
        {
            ResponseModel result = new ResponseModel();
            DocUtil docUtil = new DocUtil();
            docUtil.GenerateExcelFile(listDefUrlOrID, listDefItemId, listOrgUrlOrID, listOrgItemId, result);
            return result;
        }

        public ResponseModel GenerateExcelFilePOST(string listDefUrlOrID, string listDefItemId, string listOrgUrlOrID, string listOrgItemId)
        {
            ResponseModel result = new ResponseModel();
            DocUtil docUtil = new DocUtil();
            docUtil.GenerateExcelFile(listDefUrlOrID, listDefItemId, listOrgUrlOrID, listOrgItemId, result);
            return result;
        }

    }

}