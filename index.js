const soap = require('soap');

const url = "http://ws.prima-inform.ru/PrimaService.asmx?wsdl";

var Autentication ='<AuthHeader xmlns="http://ws.prima-inform.ru/"><Username></Username><Password></Password></AuthHeader>'

soap.createClient(url, (err, client) => {
    client.addSoapHeader(Autentication);
    // console.log(client.describe().PrimaService.PrimaServiceSoap.GetFullReport);
    // client.GetServiceInfo({}, (err, result) => {
    //     if (err) { 
    //         console.log(err)
    //     } else {
    //     console.log(result.GetServiceInfoResult.Sources.SourceInfo)
    //     }
    // })
    getOrgReport(client);
});
/**
 *
 * @param {Client} client - инстанс клиета от @function createClient 
 */
function getOrgReport (client) {
    
    client.RunRequest({ query: {
        Properties: [
            {
                PropertyValue: [
                    {
                        ProprtyId: 113,
                        Value: 7706032060
                    },
                ]
            }
        ],
        SourcesId: {
            int: [301, 303, 1001450, 1001230, 991086, 1001120, 1001550, 1001340,
                1001700, 1001650, 99106999, 718, 711, 712, 714, 713,120723, 
                991090, 120725, 991095, 120730, 715, 719, 720, 721,]
        }
    },
    }, (err, result) => {
        if (err) { 
            console.log(err)
        } else {
            // console.log(result)
            let str = myAssign({}, result);

            setTimeout(getReport, 1500, client, str);
        }
    });
}

/**
 * 
 * @param {Client} client 
 * @param {Object} str - копияе result
 */
function getReport(client, str) {
    client.GetFullReport({requestId: str.RunRequestResult.RequestId, format: 'Json'}, (err, result, rawRes, head, req) => {
        if (err) {
            console.log(err)
        } else {
            toJson(result.GetFullReportResult);
            // console.log(result.GetFullReportResult)
        }
    })
}

function myAssign(target, ...sources) {
    sources.forEach(source => {
      Object.defineProperties(target, Object.keys(source).reduce((descriptors, key) => {
        descriptors[key] = Object.getOwnPropertyDescriptor(source, key);
        return descriptors;
      }, {}));
    });
    return target;
  }

function toJson (data) {
    let dataBuffer = Buffer.from(data, 'base64');

    let text = dataBuffer.toString("utf8")

    let json = JSON.parse(text);
    console.log(json)
}