const soap = require('soap');
const config = require('./credentials.json');
const excelParser = require('./excel-parser.js')
const innValues = config.innArray;
const url = "http://ws.prima-inform.ru/PrimaService.asmx?wsdl";

const Autentication =`<AuthHeader xmlns="http://ws.prima-inform.ru/"><Username>${config.username}</Username><Password>${config.password}</Password></AuthHeader>`

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
    innValues.forEach((value, index) => {
        getOrgReport(client, value, index);
    });
});

/**
 *
 * @param {Client} client - инстанс клиета от @function createClient
 * @param {Number} innValues - ИНН
 */
function getOrgReport (client, innValue, index) {
    
        client.RunRequest({ query: {
            Properties: [
                {
                    PropertyValue: [
                        {
                            ProprtyId: 113,
                            Value: innValue
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
                getReport(client, str, index)
            }
        }, {timeout: 1000});
}

/**
 * 
 * @param {Client} client 
 * @param {Object} str - копияе result
 */

//'174c2444-7faa-4b02-9153-841e59258634'
function getReport(client, str, innIndex) {
    client.GetFullReport({requestId: str.RunRequestResult.RequestId, format: 'Json'}, (err, result, rawRes, head, req) => {
        if (err) {
            console.log(err)
        } else {
            var json = toJson(result.GetFullReportResult);
            excelParser.parseToExcel(json, innIndex);
            // console.log(result.GetFullReportResult)
        }
    }, {timeout: 1000})
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

  /**
   * 
   * @param {string} data base64 string 
   */
function toJson (data) {
    let dataBuffer = Buffer.from(data, 'base64');

    let text = dataBuffer.toString("utf8")

    let json = JSON.parse(text);
    return json
}
let sleep = () => {};