var WebAPI = {

    GetWebAPIPath: function () {
        /// <summary> Get Web API Path </summary>
        try {
            return WebAPI.GlobalContext().getClientUrl() + "/api/data/v9.1/";
        } catch (e) {

        }
    },

    GlobalContext: function () {
        var errorMessage = "Context is not available.";
        if (typeof GetGlobalContext != "undefined") {
            return GetGlobalContext();
        }
        else {
            if (typeof Xrm != "undefined") {
                return Xrm.Utility.getGlobalContext();
            }
            else { throw new Error(errorMessage); }
        }
    },

    EvalReadyState: function (state, successCallBack, errorCallBack, onComplete) {
        /// <summary> Retrieve record from CRM </summary>
        /// <param name="state" type="string">Entity Logical name</param>
        /// <param name="successCallBack" type="function">Success callback function</param>
        /// <param name="errorCallBack" type="function">Error callback function</param>
        try {
            if (state.status == 200) {
                var data = JSON.parse(state.response);
                if (typeof data["@odata.count"] !== "undefined") {
                    successCallBack(data["@odata.count"]);
                } else {
                    successCallBack(data);
                }
            } else if (state.status == 204) {
                var entityid = state.getResponseHeader("odata-entityid");
                if (entityid) {
                    entityid = entityid.substring(entityid.indexOf("(") + 1).substring(0, 36);
                    successCallBack(entityid);
                } else {
                    successCallBack();
                }
            }
            else {
                var error = JSON.parse(state.response).error;
                if (errorCallBack) {
                    errorCallBack(error);
                }
            }
            if (onComplete) {
                onComplete();
            }
        } catch (e) {

        }
    },

    Create: function (entity, object, sync, successCallBack, errorCallBack) {

        var clientURL = Xrm.Utility.getGlobalContext().getClientUrl();

        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(WebAPI.GetWebAPIPath() + entity), sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };

        req.send(JSON.stringify(object));
    },

    CreateAsUser: function (entity, object, sync, impersonateUserID, successCallBack, errorCallBack) {

        var clientURL = Xrm.Utility.getGlobalContext().getClientUrl();

        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(WebAPI.GetWebAPIPath() + entity), sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("MSCRMCallerID", impersonateUserID);
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };

        req.send(JSON.stringify(object));
    },

    Update: function (entity, id, object, sync, successCallBack, errorCallBack) {

        var req = new XMLHttpRequest();
        req.open("PATCH", encodeURI(WebAPI.GetWebAPIPath() + entity + "(" + id + ")"), sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };

        req.send(JSON.stringify(object));
    },

    UpdateAs: function (entity, id, object, sync, impersonateUserID, successCallBack, errorCallBack) {

        var req = new XMLHttpRequest();
        req.open("PATCH", encodeURI(WebAPI.GetWebAPIPath() + entity + "(" + id + ")"), sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("MSCRMCallerID", impersonateUserID);
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };

        req.send(JSON.stringify(object));
    },
    Retrieve: function (logicalName, id, columnset, sync, successCallBack, errorCallBack) {
        /// <summary> Retrieve a single record </summary>
        /// <param name="logicalName" type="string">The logical name of the record you want to retrieve (plural)</param>
        /// <param name="id" type="string">The id of the record you want to retrieve</param>
        /// <param name="columnset" type="array">Error callback function</param>
        /// <param name="successCallBack" type="function">The function to call when the entity is created.
        ///The Uri of the created entity is passed to this function.</param>
        /// <param name="errorCallBack" type="function">The function to call when there is an error. 
        ///The error will be passed to this function.</param>


        //logicalName = logicalName + "s";
        id = id.replace("{", "").replace("}", "");
        var select = "";
        var expand = "";
        var expandDictionary = [];

        if (columnset === true) {
            select = "*";
        } else {
            columnset.forEach(function (column) {
                if (column.indexOf(".") < 0) {
                    if (select.length > 0) select = select + ", ";
                    select = select + column;
                } else {
                    var columnSplit = column.split(".");
                    if (expandDictionary[columnSplit[0]]) {
                        expandDictionary[columnSplit[0]] = expandDictionary[columnSplit[0]] + ", " + columnSplit[1];
                    } else {
                        expandDictionary[columnSplit[0]] = columnSplit[1];
                    }
                }
            });
        }

        for (var key in expandDictionary) {
            if (typeof expandDictionary[key] !== 'function') {
                if (expand.length > 0) expand = expand + ", ";
                expand = expand + key + "($select=" + expandDictionary[key] + ")";
            }
        }

        var req = new XMLHttpRequest();
        var queryString = logicalName + "(" + id + ")?$select=" + select;
        if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();


    },

    RetrieveMultiple: function (logicalName, columnset, filter, sort, sync, successCallBack, errorCallBack, onComplete) {
        /// <summary> Retrieve a single record </summary>
        /// <param name="logicalName" type="string">The logical name of the record you want to retrieve (plural)</param>
        /// <param name="id" type="string">The id of the record you want to retrieve</param>
        /// <param name="columnset" type="array">Error callback function</param>
        /// <param name="successCallBack" type="function">The function to call when the entity is created.
        ///The Uri of the created entity is passed to this function.</param>
        /// <param name="errorCallBack" type="function">The function to call when there is an error. 
        ///The error will be passed to this function.</param>


        //logicalName = logicalName + "s";
        //id = id.replace("{", "").replace("}", "");
        var select = "";
        var expand = "";
        var expandDictionary = [];

        if (columnset === true) {
            select = "*";
        } else {
            columnset.forEach(function (column) {
                if (column.indexOf(".") < 0) {
                    if (select.length > 0) select = select + ", ";
                    select = select + column;
                } else {
                    var columnSplit = column.split(".");
                    if (expandDictionary[columnSplit[0]]) {
                        expandDictionary[columnSplit[0]] = expandDictionary[columnSplit[0]] + ", " + columnSplit[1];
                    } else {
                        expandDictionary[columnSplit[0]] = columnSplit[1];
                    }
                }
            });
        }

        for (var key in expandDictionary) {
            if (typeof expandDictionary[key] !== 'function') {
                if (expand.length > 0) expand = expand + ", ";
                expand = expand + key + "($select=" + expandDictionary[key] + ")";
            }
        }

        var req = new XMLHttpRequest();
        var queryString = logicalName + "?$select=" + select;
        if (expand != "") queryString = queryString + "&$expand=" + expand;
        if (filter && filter != "") queryString = queryString + "&$filter=" + filter;
        if (sort && sort != "") queryString = queryString + "&$orderby=" + sort;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, onComplete);
            }
        };
        req.send();


    },

    RetrieveMultipleAs: function (logicalName, columnset, filter, sort, sync, impersonateUserID, successCallBack, errorCallBack, onComplete) {
        /// <summary> Retrieve a single record </summary>
        /// <param name="logicalName" type="string">The logical name of the record you want to retrieve (plural)</param>
        /// <param name="id" type="string">The id of the record you want to retrieve</param>
        /// <param name="columnset" type="array">Error callback function</param>
        /// <param name="successCallBack" type="function">The function to call when the entity is created.
        ///The Uri of the created entity is passed to this function.</param>
        /// <param name="errorCallBack" type="function">The function to call when there is an error. 
        ///The error will be passed to this function.</param>


        //logicalName = logicalName + "s";
        //id = id.replace("{", "").replace("}", "");
        var select = "";
        var expand = "";
        var expandDictionary = [];

        if (columnset === true) {
            select = "*";
        } else {
            columnset.forEach(function (column) {
                if (column.indexOf(".") < 0) {
                    if (select.length > 0) select = select + ", ";
                    select = select + column;
                } else {
                    var columnSplit = column.split(".");
                    if (expandDictionary[columnSplit[0]]) {
                        expandDictionary[columnSplit[0]] = expandDictionary[columnSplit[0]] + ", " + columnSplit[1];
                    } else {
                        expandDictionary[columnSplit[0]] = columnSplit[1];
                    }
                }
            });
        }

        for (var key in expandDictionary) {
            if (typeof expandDictionary[key] !== 'function') {
                if (expand.length > 0) expand = expand + ", ";
                expand = expand + key + "($select=" + expandDictionary[key] + ")";
            }
        }

        var req = new XMLHttpRequest();
        var queryString = logicalName + "?$select=" + select;
        if (expand != "") queryString = queryString + "&$expand=" + expand;
        if (filter && filter != "") queryString = queryString + "&$filter=" + filter;
        if (sort && sort != "") queryString = queryString + "&$orderby=" + sort;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.setRequestHeader("MSCRMCallerID", impersonateUserID);
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, onComplete);
            }
        };
        req.send();


    },

    RetrieveEntityMetadataId: function (entityLogicalName, sync, successCallBack, errorCallBack) {
        var queryString = "EntityDefinitions?$select=*&$filter=LogicalName eq '" + entityLogicalName + "'";
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    RetrieveAttributeMetadataId: function (entityId, attributeLogicalName, sync, successCallBack, errorCallBack) {
        //var queryString = "EntityDefinitions?$select=*&$filter=LogicalName eq '" + entityLogicalName + "'";
        var queryString = "EntityDefinitions(" + entityId + ")/Attributes?$select=*&$filter=LogicalName eq '" + attributeLogicalName + "'";
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    RetrieveAttributesMetadata: function (entityId, sync, successCallBack, errorCallBack) {
        var queryString = "EntityDefinitions(" + entityId + ")/Attributes";
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    RetrieveAttributeMetadata: function (entityId, attributeLogicalName, sync, successCallBack, errorCallBack) {
        var queryString = "EntityDefinitions(" + entityId + ")/Attributes?$filter=LogicalName eq '" + attributeLogicalName + "'";
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    RetrieveOptionSetOptionsMetadata: function (entityId, attributeId, sync, successCallBack, errorCallBack) {
        var queryString = "EntityDefinitions(" + entityId + ")/Attributes(" + attributeId + ")/Microsoft.Dynamics.CRM.PicklistAttributeMetadata/OptionSet?$select=Options";
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();

    },

    RetrieveLocalGlobalOptionSetOptionsMetadata: function (entityName, attributeName, sync, successCallBack, errorCallBack) {
        var queryString = "EntityDefinitions(LogicalName='" + entityName + "')/Attributes(LogicalName='" + attributeName + "')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options),GlobalOptionSet($select=Options)"
        var req = new XMLHttpRequest();
        //var queryString = logicalName + "(" + id + ")?$select=" + select;
        //if (expand != "") queryString = queryString + "&$expand=" + expand;
        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();

    },

    Fetch: function (logicalName, fetchXML, sync, successCallBack, errorCallBack) {
        var req = new XMLHttpRequest();
        var queryString = logicalName + "?fetchXml=" + encodeURI(fetchXML);

        req.open("GET", WebAPI.GetWebAPIPath() + queryString, sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json;");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    CallAction: function (logicalName, entityId, actionName) {
        var req = new XMLHttpRequest();
        var queryString = logicalName + "(" + entityId.replace('{', '').replace('}', '') + ")/" + "Microsoft.Dynamics.CRM." + actionName;

        req.open("POST", WebAPI.GetWebAPIPath() + queryString);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json;");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        /*req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }getXmlValue 
        };*/
        req.send();
    },

    CallActionAs: function (logicalName, entityId, actionName, impersonateUserID,sync, successCallBack, errorCallBack) {
        var req = new XMLHttpRequest();
        var queryString = logicalName + "(" + entityId.replace('{', '').replace('}', '') + ")/" + "Microsoft.Dynamics.CRM." + actionName;

        req.open("POST", WebAPI.GetWebAPIPath() + queryString,sync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.setRequestHeader("MSCRMCallerID", impersonateUserID);

        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        req.send();
    },

    CallGlobalActionAs: function (actionName, impersonateUserID, parameters, successCallBack, errorCallBack) {
        var req = new XMLHttpRequest();
        var queryString = "/" + actionName;

        req.open("POST", WebAPI.GetWebAPIPath() + queryString);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Prefer", 'odata.include-annotations="*"');
        req.setRequestHeader("MSCRMCallerID", impersonateUserID);
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                WebAPI.EvalReadyState(this, successCallBack, errorCallBack, null);
            }
        };
        if (parameters)
            req.send(JSON.stringify(parameters));
        else
            req.send();
    },

    callAction: function (actionName, inputParams, successCallback, errorCallback, url) {
        var ns = {
            "": "http://schemas.microsoft.com/xrm/2011/Contracts/Services",
            ":s": "http://schemas.xmlsoap.org/soap/envelope/",
            ":a": "http://schemas.microsoft.com/xrm/2011/Contracts",
            ":i": "http://www.w3.org/2001/XMLSchema-instance",
            ":b": "http://schemas.datacontract.org/2004/07/System.Collections.Generic",
            ":c": "http://www.w3.org/2001/XMLSchema",
            ":d": "http://schemas.microsoft.com/xrm/2011/Contracts/Services",
            ":e": "http://schemas.microsoft.com/2003/10/Serialization/",
            ":f": "http://schemas.microsoft.com/2003/10/Serialization/Arrays",
            ":g": "http://schemas.microsoft.com/crm/2011/Contracts",
            ":h": "http://schemas.microsoft.com/xrm/2011/Metadata",
            ":j": "http://schemas.microsoft.com/xrm/2011/Metadata/Query",
            ":k": "http://schemas.microsoft.com/xrm/2013/Metadata",
            ":l": "http://schemas.microsoft.com/xrm/2012/Contracts",
            //":c": "http://schemas.microsoft.com/2003/10/Serialization/" // Conflicting namespace for guid... hardcoding in the _getXmlValue bit
        };
        var requestXml = "<s:Envelope";
        // Add all the namespaces
        for (var i in ns) {
            requestXml += " xmlns" + i + "='" + ns[i] + "'";
        }
        requestXml += ">" +
              "<s:Body>" +
                "<Execute>" +
                  "<request>";

        if (inputParams != null && inputParams.length > 0) {
            requestXml += "<a:Parameters>";
            // Add each input param
            for (var i = 0; i < inputParams.length; i++) {
                var param = inputParams[i];

                var value = WebAPI._getXmlValue(param.key, param.type, param.value);

                requestXml += value;
            }
            requestXml += "</a:Parameters>";
        }
        else {
            requestXml += "<a:Parameters />";
        }

        requestXml += "<a:RequestId i:nil='true' />" +
                    "<a:RequestName>" + actionName + "</a:RequestName>" +
                  "</request>" +
                "</Execute>" +
              "</s:Body>" +
            "</s:Envelope>";

        WebAPI._callActionBase(requestXml, successCallback, errorCallback, url);
    },
    _callActionBase: function (requestXml, successCallback, errorCallback, url) {
        if (url == null) {
            url = Xrm.Utility.getGlobalContext().getClientUrl();
        }

        var req = new XMLHttpRequest();
        req.open("POST", url + "/api/data/v9.1/web", false);
        req.setRequestHeader("Accept", "application/xml, text/xml, */*");
        req.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
        req.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute");

        req.onreadystatechange = function () {
            if (req.readyState == 4) {
                if (req.status == 200) {
                    // If there's no successCallback we don't need to check the outputParams
                    if (successCallback) {
                        // Yucky but don't want to risk there being multiple 'Results' nodes or something
                        var resultsNode = req.responseXML.childNodes[0].childNodes[0].childNodes[0].childNodes[0].childNodes[1]; // <a:Results>

                        // Action completed successfully - get output params
                        var responseParams = WebAPI._getChildNodes(resultsNode, "a:KeyValuePairOfstringanyType");
                        if (responseParams != null) {
                            var outputParams = {};
                            for (i = 0; i < responseParams.length; i++) {
                                var attrNameNode = WebAPI._getChildNode(responseParams[i], "b:key");
                                var attrValueNode = WebAPI._getChildNode(responseParams[i], "b:value");

                                var attributeName = WebAPI._getNodeTextValue(attrNameNode);
                                var attributeValue = WebAPI._getValue(attrValueNode);

                                // v1.0 - Deprecated method using key/value pair and standard array
                                //outputParams.push({ key: attributeName, value: attributeValue.value });

                                // v2.0 - Allows accessing output params directly: outputParams["Target"].attributes["new_fieldname"];
                                outputParams[attributeName] = attributeValue.value;

                                /*
                                RETURN TYPES:
                                    DateTime = Users local time (JavaScript date)
                                    bool = true or false (JavaScript boolean)
                                    OptionSet, int, decimal, float, etc = 1 (JavaScript number)
                                    guid = string
                                    EntityReference = { id: "guid", name: "name", entityType: "account" }
                                    Entity = { logicalName: "account", id: "guid", attributes: {}, formattedValues: {} }
                                    EntityCollection = [{ logicalName: "account", id: "guid", attributes: {}, formattedValues: {} }]
            
                                Attributes for entity accessed like: entity.attributes["new_fieldname"].value
                                For entityreference: entity.attributes["new_fieldname"].value.id
                                Make sure attributes["new_fieldname"] is not null before using .value
                                Or use the extension method entity.get("new_fieldname") to get the .value
                                Also use entity.formattedValues["new_fieldname"] to get the string value of optionsetvalues, bools, moneys, etc
                                */
                            }
                        }

                        // Make sure the callback accepts exactly 1 argument - use dynamic function if you want more
                        successCallback(outputParams);
                    }
                }
                else {
                    // Error has occured, action failed
                    if (errorCallback) {
                        var message = null;
                        var traceText = null;
                        try {
                            message = WebAPI._getNodeTextValueNotNull(req.responseXML.getElementsByTagName("Message"));
                            traceText = WebAPI._getNodeTextValueNotNull(req.responseXML.getElementsByTagName("TraceText"));
                        } catch (e) { }
                        if (message == null) { message = "Error executing Action. Check input parameters or contact your CRM Administrator"; }
                        errorCallback(message, traceText);
                    }
                }
            }
        };

        req.send(requestXml);
    },

    _getXmlValue: function (key, dataType, value) {
        var xml = "";
        var xmlValue = "";

        var extraNamespace = "";

        // Check the param type to determine how the value is formed
        switch (dataType) {
            case WebAPI.Type.String:
                xmlValue = WebAPI._htmlEncode(value) || ""; // Allows fetchXml strings etc
                break;
            case WebAPI.Type.DateTime:
                xmlValue = value.toISOString() || "";
                break;
            case WebAPI.Type.EntityReference:
                xmlValue = "<a:Id>" + (value.id || "") + "</a:Id>" +
                      "<a:LogicalName>" + (value.entityType || "") + "</a:LogicalName>" +
                      "<a:Name i:nil='true' />";
                break;
            case WebAPI.Type.OptionSet:
            case WebAPI.Type.Money:
                xmlValue = "<a:Value>" + (value || 0) + "</a:Value>";
                break;
            case WebAPI.Type.Entity:
                xmlValue = WebAPI._getXmlEntityData(value);
                break;
            case WebAPI.Type.EntityCollection:
                if (value != null && value.length > 0) {
                    var entityCollection = "";
                    for (var i = 0; i < value.length; i++) {
                        var entityData = WebAPI._getXmlEntityData(value[i]);
                        if (entityData !== null) {
                            entityCollection += "<a:Entity>" + entityData + "</a:Entity>";
                        }
                    }
                    if (entityCollection !== null && entityCollection !== "") {
                        xmlValue = "<a:Entities>" + entityCollection + "</a:Entities>" +
                            "<a:EntityName i:nil='true' />" +
                            "<a:MinActiveRowVersion i:nil='true' />" +
                            "<a:MoreRecords>false</a:MoreRecords>" +
                            "<a:PagingCookie i:nil='true' />" +
                            "<a:TotalRecordCount>0</a:TotalRecordCount>" +
                            "<a:TotalRecordCountLimitExceeded>false</a:TotalRecordCountLimitExceeded>";
                    }
                }
                break;
            case WebAPI.Type.Guid:
                // I don't think guid fields can even be null?
                xmlValue = value || WebAPI._emptyGuid;

                // This is a hacky fix to get guids working since they have a conflicting namespace :(
                extraNamespace = " xmlns:c='http://schemas.microsoft.com/2003/10/Serialization/'";
                break;
            default: // bool, int, double, decimal
                xmlValue = value != undefined ? value : null;
                break;
        }

        xml = "<a:KeyValuePairOfstringanyType>" +
                "<b:key>" + key + "</b:key>" +
                "<b:value i:type='" + dataType + "'" + extraNamespace;

        // nulls crash if you have a non-self-closing tag
        if (xmlValue === null || xmlValue === "") {
            xml += " i:nil='true' />";
        }
        else {
            xml += ">" + xmlValue + "</b:value>";
        }

        xml += "</a:KeyValuePairOfstringanyType>";

        return xml;
    },
    _htmlEncode: function (s) {
        if (typeof s !== "string") { return s; }

        return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    },
    _getValue: function (node) {
        var value = null;
        var type = null;

        if (node != null) {
            type = node.getAttribute("i:type") || node.getAttribute("type");

            // If the parameter/attribute is null, there won't be a type either
            if (type != null) {
                // Get the part after the ':' (since Chrome doesn't have the ':')
                var valueType = type.substring(type.indexOf(":") + 1).toLowerCase();

                if (valueType == "entityreference") {
                    // Gets the lookup object
                    var attrValueIdNode = WebAPI._getChildNode(node, "a:Id");
                    var attrValueEntityNode = WebAPI._getChildNode(node, "a:LogicalName");
                    var attrValueNameNode = WebAPI._getChildNode(node, "a:Name");

                    var lookupId = WebAPI._getNodeTextValue(attrValueIdNode);
                    var lookupName = WebAPI._getNodeTextValue(attrValueNameNode);
                    var lookupEntity = WebAPI._getNodeTextValue(attrValueEntityNode);

                    value = new WebAPI.EntityReference(lookupEntity, lookupId, lookupName);
                }
                else if (valueType == "entity") {
                    // Gets the entity data, and all attributes
                    value = WebAPI._getEntityData(node);
                }
                else if (valueType == "entitycollection") {
                    // Loop through each entity, returns each entity, and all attributes
                    var entitiesNode = WebAPI._getChildNode(node, "a:Entities");
                    var entityNodes = WebAPI._getChildNodes(entitiesNode, "a:Entity");

                    value = [];
                    if (entityNodes != null && entityNodes.length > 0) {
                        for (var i = 0; i < entityNodes.length; i++) {
                            value.push(WebAPI._getEntityData(entityNodes[i]));
                        }
                    }
                }
                else if (valueType == "aliasedvalue") {
                    // Gets the actual data type of the aliased value
                    // Key for these is "alias.fieldname"
                    var aliasedValue = WebAPI._getValue(WebAPI._getChildNode(node, "a:Value"));
                    if (aliasedValue != null) {
                        value = aliasedValue.value;
                        type = aliasedValue.type;
                    }
                }
                else {
                    // Standard fields like string, int, date, money, optionset, float, bool, decimal
                    // Output will be string, even for number fields etc
                    var stringValue = WebAPI._getNodeTextValue(node);

                    if (stringValue != null) {
                        switch (valueType) {
                            case "datetime":
                                value = new Date(stringValue);
                                break;
                            case "int":
                            case "money":
                            case "optionsetvalue":
                            case "double": // float
                            case "decimal":
                                value = Number(stringValue);
                                break;
                            case "boolean":
                                value = stringValue.toLowerCase() === "true";
                                break;
                            default:
                                value = stringValue;
                        }
                    }
                }
            }
        }

        return new WebAPI.Attribute(value, type);
    },
    _getChildNodes: function (node, childNodesName) {
        var childNodes = [];

        for (var i = 0; i < node.childNodes.length; i++) {
            if (node.childNodes[i].tagName == childNodesName) {
                childNodes.push(node.childNodes[i]);
            }
        }
        return childNodes;
    },
    _getChildNode: function (node, childNodeName) {
        var nodes = WebAPI._getChildNodes(node, childNodeName);

        if (nodes != null && nodes.length > 0) { return nodes[0]; }
        else { return null; }
    },

    _getNodeTextValueNotNull: function (nodes) {
        var value = "";

        for (var i = 0; i < nodes.length; i++) {
            if (value === "") {
                value = WebAPI._getNodeTextValue(nodes[i]);
            }
        }

        return value;
    },
    EntityReference: function (entityType, id, name) {
        this.id = id || "00000000-0000-0000-0000-000000000000";
        this.name = name || "";
        this.entityType = entityType || "";
    },
    _emptyGuid: "00000000-0000-0000-0000-000000000000",


    Type: {
        Bool: "c:boolean",
        Float: "c:double", // Not a typo
        Decimal: "c:decimal",
        Int: "c:int",
        String: "c:string",
        DateTime: "c:dateTime",
        Guid: "c:guid",
        EntityReference: "a:EntityReference",
        OptionSet: "a:OptionSetValue",
        Money: "a:Money",
        Entity: "a:Entity",
        EntityCollection: "a:EntityCollection"
    },

    _getNodeTextValue: function (node) {
        if (node != null) {
            var textNode = node.firstChild;
            if (textNode != null) {
                return textNode.textContent || textNode.nodeValue || textNode.data || textNode.text;
            }
        }

        return "";
    },
    Attribute: function (value, type) {
        this.value = value != undefined ? value : null;
        this.type = type || "";
    }


}

