// JavaScript source code

var Case = window.Case || {};

(function () {
    // Define some global variables
    var priorityFlag = null;
    // Code to run in the form OnLoad event
    this.FormOnLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();
        Case.FilterProjects(executionContext);
        Case.SoftwareDetails(executionContext);
        Case.ShowHideFieldServiceTab(executionContext);
        Case.FilterServiceContract(executionContext);
        Case.PopulateContractOnLoad(executionContext);
        Case.editCreditCheck(executionContext);
    };
    // Code to run in the form OnSave event 
    this.FormOnSave = function (executionContext) {
        var formContext = executionContext.getFormContext();
    };


    //Show Software details, if call type is software or software remote work
    this.SoftwareDetails = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var softwareTab = _formContext.ui.tabs.get("Software");
        var callType = _formContext.getAttribute("infy_calltype").getValue();
        if (callType !== null) {
            if (callType[0].name === "Software" || callType[0].name === "Software Remote Work")
                softwareTab.setVisible(true);
            else
                softwareTab.setVisible(false);
        }
    };

    //Show/Hide Field Service tab, if credit check is passed
    this.ShowHideFieldServiceTab = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var fieldServiceTab = _formContext.ui.tabs.get("FieldService");
        var creditCheck = _formContext.getAttribute("infy_creditcheckstatus").getValue();
        if (creditCheck === EnumCreditCheck.Passed) { //Passed
            fieldServiceTab.setVisible(true);
        }
        else {
            fieldServiceTab.setVisible(false);
        }
    };

    // Empty Departments based on LOB
    this.EmptyDepartment = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        _formContext.getAttribute("infy_departments").setValue(null);
    }

    // Clear Asset fields on customer change
    this.ClearFieldsCustomerChange = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        _formContext.getAttribute("infy_customerasset").setValue(null);
        _formContext.getAttribute("infy_serial").setValue(null);
        _formContext.getAttribute("infy_equip").setValue(null);
        _formContext.getAttribute("infy_model").setValue(null);
        _formContext.getAttribute("infy_equipmentcontact").setValue(null);
        _formContext.getAttribute("infy_contract").setValue(null);
        _formContext.getAttribute("infy_contractno2").setValue(null);

    };

    // Filter Projects based on Project type
    this.FilterProjects = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        _formContext.getControl("infy_project").addPreSearch(Case.ApplyFilterProject);
    };

    //Filter Projects lookup based on Project type selected
    this.ApplyFilterProject = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var projectType = _formContext.getAttribute("infy_projecttype").getValue();
        if (projectType !== "") {
            var fetchXml = "<filter type='and'><condition attribute='infy_projecttype' operator='eq' value='" + projectType + "' /></filter>";
            _formContext.getControl("infy_project").addCustomFilter(fetchXml);
        }
    };

    // Filter service account based on customer
    this.FilterServiceContract = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var serviceAccount = _formContext.getAttribute("customerid").getValue();
        if (serviceAccount !== null) {
            _formContext.getControl("infy_contract").addPreSearch(Case.ApplyFilterServiceContract);
        }
    };

    //Filter service account based on customer FetchXML
    this.ApplyFilterServiceContract = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var serviceAccount = _formContext.getAttribute("customerid").getValue();
        var serviceAccountID = serviceAccount[0].id.replace('{', '').replace('}', '');
        var entityLogicalName = "/accounts(" + serviceAccountID + ")";
        var columnsToRetrieveAccount = "_msdyn_billingaccount_value,";
        Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieveAccount + "")
            .then(function (request) {
                if (request !== null) {
                    results = JSON.parse(request.response);
                    var BillingAccount = results["_msdyn_billingaccount_value"];
                    var customer = BillingAccount || serviceAccountID;
                    var fetchXml = "<filter type='and'><condition attribute='customerid' operator='eq' value='" + customer + "' /></filter>";
                    _formContext.getControl("infy_contract").addCustomFilter(fetchXml);
                }
            });
    };

    //fetch serial,equip,model and eqip. contact from customer asset on selection of asset lookup field
    this.RetrieveAssetInfo = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var customerAsset = _formContext.getAttribute("infy_customerasset").getValue();
        if (customerAsset !== null) {
            var entityId = customerAsset[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/msdyn_customerassets(" + entityId + ")";
            var columnsToRetrieve = "infy_serialno,infy_equipno,infy_model,_infy_equipmentcontact_value";
            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                .then(function (request) {
                    if (request !== null) {
                        results = JSON.parse(request.response);
                        _formContext.getAttribute("infy_serial").setValue(results["infy_serialno"]);
                        _formContext.getAttribute("infy_equip").setValue(results["infy_equipno"]);
                        _formContext.getAttribute("infy_model").setValue(results["infy_model"]);
                        var equipContactId = results["_infy_equipmentcontact_value"] || "";
                        if (equipContactId !== "") {
                            var equipContactValue = results["_infy_equipmentcontact_value@OData.Community.Display.V1.FormattedValue"];
                            if (equipContactId !== null || equipContactId !== undefined) {
                                var EquipmentContact = new Array();
                                EquipmentContact[0] = new Object();
                                EquipmentContact[0].id = equipContactId;
                                EquipmentContact[0].name = equipContactValue;
                                EquipmentContact[0].entityType = "contact";
                                _formContext.getAttribute("infy_equipmentcontact").setValue(EquipmentContact);

                            }
                        }
                    }
                })
                .catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);
                });
        }
        else {
            _formContext.getAttribute("infy_serial").setValue(null);
            _formContext.getAttribute("infy_equip").setValue(null);
            _formContext.getAttribute("infy_model").setValue(null);
            _formContext.getAttribute("infy_equipmentcontact").setValue(null);
        }
    };

   //fetch Contract based on customer
    this.RetrieveContractCustomer = function (executionContext,OnLoad) {
        var _formContext = executionContext.getFormContext();
        var serviceAccount = _formContext.getAttribute("customerid").getValue();
        if (serviceAccount !== null && serviceAccount[0].entityType == "account") {
            var serviceAccountID = serviceAccount[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/accounts(" + serviceAccountID + ")";
            var columnsToRetrieveAccount = "_msdyn_billingaccount_value";
            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieveAccount + "")
                .then(function (request) {
                    if (request !== null) {
                        results = JSON.parse(request.response);
                        BillingAccount = results["_msdyn_billingaccount_value"];
                        customer = BillingAccount || serviceAccountID;
                        var entityLogicalName = "/entitlements";
                        var fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
                            "<entity name='entitlement'>" +
                            "<attribute name='name' />" +
                            "<attribute name='entitlementid' />" +
                            "<attribute name='infy_contractnumber' />" +
                            "<attribute name='msdyn_entitlementprioritization' />" +
                            "<order attribute='msdyn_entitlementprioritization' descending='false' />" +
                            "<order attribute='enddate' descending='false' />" +
                            "<filter type='and'>" +
                            "<condition attribute='customerid' operator='eq' value='" + customer + "' />" +
                            "<condition attribute='statecode' operator='eq' value='1' />" +
                            "<condition attribute='entitytype' operator='eq' value='192350000' />" +
                            "</filter>" +
                            "<link-entity name='msdyn_entitlementapplication' from='msdyn_entitlement' to='entitlementid' link-type='outer' alias='ae' />" +
                            "<filter type='and'>" +
                            "<condition entityname='ae' attribute='msdyn_entitlement' operator='null' />" +
                            "</filter>" +
                            "</entity>" +
                            "</fetch>";
                        Sdk.request("GET", entityLogicalName + "?fetchXml=" + fetchXML)
                            .then(function (request) {
                                if (request !== null) {
                                    results = JSON.parse(request.response);
                                    if (results.value.length > 0) {
                                        var priorityContractlist = results.value.filter(value => value["msdyn_entitlementprioritization"] !== undefined);
                                        var finalContract = null;
                                        if (priorityContractlist.length > 0) {
                                            finalContract = priorityContractlist[0];
                                            priorityFlag = finalContract["msdyn_entitlementprioritization"];
                                        }
                                        else
                                            finalContract = results.value[0];
                                        if (finalContract !== null && OnLoad == null) {
                                            var contractId = finalContract["entitlementid"] || "";
                                            if (contractId !== "") { 
                                                var contractValue = finalContract["name"];
                                                if (contractValue !== null || contractValue !== undefined) {
                                                    var contract = new Array();
                                                    contract[0] = new Object();
                                                    contract[0].id = contractId;
                                                    contract[0].name = contractValue;
                                                    contract[0].entityType = "entitlement";
                                                    _formContext.getAttribute("infy_contract").setValue(contract);
                                                    var contractNumber = finalContract["infy_contractnumber"];
                                                    _formContext.getAttribute("infy_contractno2").setValue(contractNumber);
                                                }
                                            }
                                        }
                                    }
                                }
                            });
                    }
                }).catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);
                });
        }
        else {
            _formContext.getAttribute("infy_contractnumber").setValue(null);
            _formContext.getAttribute("infy_contract").setValue(null);
        }
    };
       
    //fetch Contract based on asset
    this.RetrieveContract = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var BillingAccount = null, customer = null, assetCategory = null, customerAsset = null;
        //var  = _formContext.getAttribute("infy_customerasset").getValue();
        var customerAsset = _formContext.getAttribute("infy_customerasset").getValue();
        var serviceAccount = _formContext.getAttribute("customerid").getValue();
        if (serviceAccount != null) {
            var serviceAccountID = serviceAccount[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/accounts(" + serviceAccountID + ")";
            var columnsToRetrieveAccount = "_msdyn_billingaccount_value";
            if (customerAsset !== null && serviceAccount[0].entityType == "account") {
                Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieveAccount + "")
                    .then(function (request) {
                        if (request !== null) {
                            results = JSON.parse(request.response);
                            BillingAccount = results["_msdyn_billingaccount_value"];
                            customer = BillingAccount || serviceAccountID;
                            var customerAssetID = customerAsset[0].id.replace('{', '').replace('}', '');
                            var entityLogicalName = "/msdyn_customerassets(" + customerAssetID + ")";
                            var columnsToRetrieveAsset = "_msdyn_customerassetcategory_value";
                            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieveAsset + "")
                                .then(function (request) {
                                    if (request !== null) {
                                        results = JSON.parse(request.response);
                                        assetCategory = results["_msdyn_customerassetcategory_value"];
                                        if (customerAsset !== null && customer !== null) {
                                            var customerAssetId = customerAsset[0].id.replace('{', '').replace('}', '');
                                            var entityName = "/msdyn_entitlementapplications";
                                            var fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                                                "<entity name='msdyn_entitlementapplication' >" +
                                                "<attribute name='msdyn_name' />" +
                                                "<attribute name='msdyn_serviceaccount' />" +
                                                "<attribute name='msdyn_entitlement' />" +
                                                "<attribute name='msdyn_customerasset' />" +
                                                "<attribute name='msdyn_assetcategory' />" +
                                                "<filter type='or'>" +
                                                "<condition attribute='msdyn_customerasset' operator='eq' value='" + customerAssetId + "' />";
                                            if (assetCategory !== null) {
                                                fetchXML = fetchXML + "<condition attribute='msdyn_assetcategory' operator='eq' value='" + assetCategory + "' />";
                                            }
                                            fetchXML = fetchXML + "</filter>" +
                                                "<link-entity name='entitlement' from='entitlementid' to='msdyn_entitlement' link-type='inner' alias='msdyn_entitlement'>" +
                                                "<attribute name='infy_contractnumber' />" +
                                                "<attribute name='msdyn_pricelisttoapply' />" +
                                                "<attribute name='msdyn_entitlementprioritization' />" +
                                                "<order attribute='msdyn_entitlementprioritization' descending='false' />" +
                                                "<order attribute='enddate' descending='true' />" +
                                                "<filter type='and'>" +
                                                "<condition attribute='customerid' operator='eq' value='" + customer + "' />" +
                                                "<condition attribute='statecode' operator='eq' value='1' />" +
                                                "</filter></link-entity></entity></fetch>";
                                            Sdk.request("GET", entityName + "?fetchXml=" + fetchXML)
                                                .then(function (request) {
                                                    if (request !== null) {
                                                        results = JSON.parse(request.response);
                                                        if (results.value.length > 0) {
                                                            var priorityContractlist = results.value.filter(value => value["msdyn_entitlement.msdyn_entitlementprioritization"] !== undefined);
                                                            var finalContract = null; priority = null;
                                                            if (priorityContractlist.length > 0) {
                                                                finalContract = priorityContractlist[0];
                                                                priority = finalContract["msdyn_entitlement.msdyn_entitlementprioritization"];
                                                            }
                                                            else
                                                                finalContract = results.value[0];
                                                            if (priorityFlag !== null && priority === null)
                                                                return;
                                                            if (finalContract !== null && (priorityFlag === null || (priority < priorityFlag))) {
                                                                var contractId = finalContract["_msdyn_entitlement_value"] || "";
                                                                if (contractId !== "") {
                                                                    var contractValue = finalContract["_msdyn_entitlement_value@OData.Community.Display.V1.FormattedValue"];
                                                                    if (contractValue !== null || contractValue !== undefined) {
                                                                        var contract = new Array();
                                                                        contract[0] = new Object();
                                                                        contract[0].id = contractId;
                                                                        contract[0].name = contractValue;
                                                                        contract[0].entityType = "entitlement";
                                                                        var contractNumber = finalContract["msdyn_entitlement.infy_contractnumber"];
                                                                        _formContext.getAttribute("infy_contract").setValue(contract);
                                                                        _formContext.getAttribute("infy_contractno2").setValue(contractNumber);
                                                                        _formContext.data.save();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                });
                                        }
                                    }
                                });
                        }
                    }).catch(function (err) {
                        Xrm.Utility.alertDialog(err.message);
                    });
            }
            else {
                _formContext.getAttribute("infy_contractnumber").setValue(null);
                _formContext.getAttribute("infy_contract").setValue(null);
                Case.RetrieveContractCustomer(executionContext);
            }
        }

    };

    //Papulate contract onload
    this.PopulateContractOnLoad = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var serviceAccount = _formContext.getAttribute("customerid").getValue();
        var asset = _formContext.getAttribute("infy_customerasset").getValue();
        var contract = _formContext.getAttribute("infy_contract").getValue();
        if (serviceAccount !== null && contract == null) {
            Case.RetrieveContractCustomer(executionContext);
        }
        else if (serviceAccount !== null && contract != null && asset == null) {
            Case.RetrieveContractCustomer(executionContext, "OnLoad");
        }
    };

    //fetch contract number from contract entity, when user selects Support contract  lookup
    this.RetrieveContractInfo = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var contract = _formContext.getAttribute("entitlementid").getValue();
        if (contract !== null) {
            entityId = contract[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/entitlements(" + entityId + ")";
            var columnsToRetrieve = "infy_contractnumber";
            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                .then(function (request) {
                    if (request !== null) {
                        results = JSON.parse(request.response);
                        _formContext.getAttribute("infy_contractnumber").setValue(results["infy_contractnumber"]);
                    }
                })
                .catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);
                });
        }
        else {
            _formContext.getAttribute("infy_contractnumber").setValue(null);
            _formContext.getAttribute("infy_expirationcopies").setValue(null);
        }
    };

    //fetch contract number from contract entity, when user selects Support contract  lookup
    this.RetrieveServiceContractInfo = function (executionContext) {
        var _formContext = executionContext.getFormContext();
        var contract = _formContext.getAttribute("infy_contract").getValue();
        if (contract !== null) {
            entityId = contract[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/entitlements(" + entityId + ")";
            var columnsToRetrieve = "infy_contractnumber";
            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                .then(function (request) {
                    if (request !== null) {
                        results = JSON.parse(request.response);
                        _formContext.getAttribute("infy_contractno2").setValue(results["infy_contractnumber"]);
                    }
                })
                .catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);
                });
        }
        else {
            _formContext.getAttribute("infy_contractno2").setValue(null);
            _formContext.getAttribute("infy_expirationcopies2").setValue(null);
        }
    };


    this.CreditCheck = function (executionContext) {
        try {
           
            var _formContext = executionContext;
            debugger;
            //Case.ForceRefresh(_formContext);
            var parameters = {};
            var customerId = _formContext.getAttribute("customerid");
            if (customerId !== null && customerId !== undefined) {
                var accountorcontact = customerId.getValue();
                if (accountorcontact !== null && accountorcontact !== undefined) {
                    var accountorcontactEntityName = accountorcontact[0].entityType;
                    if (accountorcontactEntityName !== null && accountorcontactEntityName !== undefined) {
                        if (accountorcontactEntityName === "account") {
                            Xrm.Utility.showProgressIndicator("Credit check process inprogress..please wait");
                            parameters._accountId = accountorcontact[0].id.replace('{', '').replace('}', '');
                            if (parameters._accountId !== null && parameters._accountId !== undefined) {
                                Sdk.request("POST", "/infy_CreditCheckOnDemand", parameters)
                                    .then(function (request) {
                                        if (request !== null) {
                                            var result = JSON.parse(request.response);
                                            if (result !== undefined && result["_creditcheck"] !== undefined) {
                                                var creditcheck = JSON.parse(result["_creditcheck"]);
                                                Case.UpdateCase(_formContext, creditcheck);
                                            }
                                        }
                                    })
                                    .catch(function (err) {
                                        Xrm.Utility.alertDialog(err.message);

                                    });
                            } else {
                                Xrm.Utility.alertDialog("Credit check process not completed");
                            }

                        } else {
                            Xrm.Utility.alertDialog("Please select Account as Customer on Case form, if you are trying for credit check or converting Case to Work Order");
                        }
                    }
                }
            }

        } catch (error) {
            console.log("CreditCheck :" + error.message);
        } finally {
            setTimeout(function () { Xrm.Utility.closeProgressIndicator(); }, 1000);
        }
    };
    //update the Case Credit Check Status based on Credit Check values 
    this.UpdateCase = function (_formContext, creditcheck) {
        try {
            var caseEntity = {};
            if (creditcheck !== null & creditcheck.length > 0) {
                var creditCheckAmount = parseFloat(creditcheck[1]) - parseFloat(creditcheck[0]);
                if (creditCheckAmount > 0) {
                    caseEntity.infy_creditcheckstatus = EnumCreditCheck.Passed;//passed
                }
                else if (creditCheckAmount < 0) {
                    caseEntity.infy_creditcheckstatus = EnumCreditCheck.Falied;//failed
                }
                var recordId = _formContext.data.entity.getId().replace('{', '').replace('}', '');
                var entityLogicalName = "/incidents(" + recordId + ")";
                //Updating the Case record  
                Sdk.request("PATCH", entityLogicalName, caseEntity).then(function (request) {
                    if (request !== null) {
                        if (request.status === 204) {
                            console.log("Case Updated successfully");
                            Xrm.Utility.showProgressIndicator("Credit check process completed");
                            //Xrm.Page.data.refresh(true);
                            _formContext.data.refresh();
                            Case.ForceRefresh(_formContext);
                            //_formContext.ui.refreshRibbon();
                            //window.location.reload();

                        }

                    }
                }).catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);

                });
            } else {
                Xrm.Utility.alertDialog("Credit Limit information not found for this account");
            }

        } catch (error) {
            console.log("UpdateCase :" + error.message);
        } finally {
            setTimeout(function () {


                Xrm.Utility.closeProgressIndicator();

            }, 1000);
        }
        
    };

    const EnumCreditCheck = {
        Passed: 854360000,
        Falied: 854360001,
        Pending: 854360002
    };
    this.EnableWorkorder = function (executionContext) {

        var _formContext = executionContext;
        var enableWorkorder = false;
        if (_formContext.getAttribute("infy_creditcheckstatus").getValue() === EnumCreditCheck.Passed) {
            enableWorkorder = true;
        }
        return enableWorkorder;
    };
    this.ForceRefresh = function (_formContext) {
        var entityFormOptions = {};
        entityFormOptions["entityName"] = "incident";
        entityFormOptions["entityId"] = _formContext.data.entity.getId();

        // Open the form.
        Xrm.Navigation.openForm(entityFormOptions).then(
            function (success) {
                console.log(success);
            },
            function (error) {
                console.log(error);
            });
    };
    //Allow to Edit  the Credit Check Status field  only for the Specified User Roles 
    this.editCreditCheck = function (executionContext) {
    
      var roles = Xrm.Utility.getGlobalContext().userSettings.roles.getAll();
        var _formContext = executionContext.getFormContext();
        var creditCheckDisable = true;
        for (var i = 0; i < roles.length; i++) {
            if (roles[i].name === "System Administrator" || roles[i].name === "System Customizer" || roles[i].name === "CSR Manager") {
                creditCheckDisable = false;
                break;
            }
        }
        _formContext.getControl("infy_creditcheckstatus").setDisabled(creditCheckDisable);
    };



}).call(Case);