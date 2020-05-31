using System.ServiceModel;
using System.ServiceModel.Web;

namespace Sibur.SharePoint.DefList
{
    [ServiceContract]
    public interface IDocService
    {

        [OperationContract]
        [WebInvoke(Method = "GET",
            ResponseFormat = WebMessageFormat.Xml,
            BodyStyle = WebMessageBodyStyle.Wrapped,
            UriTemplate = "DocServiceCall({SampleValue})")]
        string DocServiceCall(string SampleValue);
        

        [OperationContract]
        [WebGet(UriTemplate = "GenerateExcelFileGET/{listDefUrlOrID}/{listDefItemId}/{listOrgUrlOrID}/{listOrgItemId}",
        ResponseFormat = WebMessageFormat.Json)]
        ResponseModel GenerateExcelFileGET(string listDefUrlOrID, string listDefItemId, string listOrgUrlOrID, string listOrgItemId);



        [OperationContract]
        [WebInvoke(BodyStyle = WebMessageBodyStyle.Wrapped,
            Method = "POST",
            RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        ResponseModel GenerateExcelFilePOST(string listDefUrlOrID, string listDefItemId, string listOrgUrlOrID, string listOrgItemId);

    }

}


