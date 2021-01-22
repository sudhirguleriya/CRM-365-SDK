var QuoteLine = window.QuoteLine || {};

(function () {

    this.FormOnLoad = function (exeContext) {
         QuoteLine.CheckStatus(exeContext);
        QuoteLine.ShowPeriodByProjectType(exeContext);
        QuoteLine.FilterParentProducts(exeContext);
       
    };

    this.FormOnSave = function (exeContext) {
        QuoteLine.OnQuoteLineProductCheck(exeContext);
    };

    this.CheckOnhandInventory = function (productNumber, businessUnitId, warehouseId, productName, quantity, counter) {        
        var productInfo = "";
        var parameters = {};
        parameters.BusinessUnitName = businessUnitId.toLowerCase();
        parameters.ProductNumber = productNumber;
        parameters.WarehouseID = warehouseId;

        var req = new XMLHttpRequest();
        req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/infy_CheckOnhandInventory", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                  //  debugger;
                    var results = JSON.parse(this.response);
                    var onHandQuantity = 0;
                    if (JSON.parse(results.Result)["value"]["length"] != 0) {
                        onHandQuantity = (JSON.parse(results.Result)["value"][0])["OnHandQuantity"];
                        if (onHandQuantity <= 0) {
                            productInfo = counter + "). " + productName + " - (Requested : " + quantity + ", Available : " + onHandQuantity + ")"
                        }
                        else if (quantity > onHandQuantity)//requested quantity is not fully available on F&O
                            productinfo = counter + "). " + productName + "-(Requested:" + quantity + ", Available:" + onHandQuantity + "). "
                    }
                    else {
                       
                        productInfo = counter + "). " + productName + " - (Requested : " + quantity + ", Available : " + onHandQuantity + ")"
                    }
                }
            }
        };
        req.send(JSON.stringify(parameters));
        return productInfo;
    };

    this.OnQuoteLineProductCheck = function (exeContext) {
     //   debugger;
        var _formContext = exeContext.getFormContext();
        var productMessages = [];
        var counter = 0;
        var productId = null;
        var quoteId = _formContext.getAttribute("quoteid");
        var quantity = _formContext.getAttribute("quantity").getValue() == null ? 0 : _formContext.getAttribute("quantity").getValue();
        if (quoteId != null && quoteId.getValue() != null) {
            quoteId = quoteId.getValue()[0].id.replace('{', '').replace('}', '');
            if (quoteId != null) {
                var filter = "&$filter=quoteid eq " + quoteId + "";
                var query = "?$select=statecode&$expand=owningbusinessunit($select=name),infy_Warehouses($select=infy_warehousenumber)" + filter;
                Sdk.request("GET", "/quotes" + query)
                    .then(function (request) {
                        if (request != null) {
                            var results = JSON.parse(request.response);//getting result 
                            var businessUnitId;
                            var warehouseId;
                            if (results.value[0].statecode == 0) {
                                if (results.value[0].owningbusinessunit != null)
                                    businessUnitId = results.value[0].owningbusinessunit["name"];
                                if (results.value[0].infy_Warehouses != null) {
                                    warehouseId = results.value[0].infy_Warehouses["infy_warehousenumber"];
                                }
                            }
                            return { businessUnitId, warehouseId }

                        }
                    })
                    .then(function (request) {
                        if (request != undefined) {
                           // debugger;
                            var businessUnit = request.businessUnitId;
                            var warehouseNumber = request.warehouseId;
                            var quoteLineProductId = _formContext.getAttribute("productid").getValue();
                           
                            if (quoteLineProductId != null){
                               var productName = quoteLineProductId[0].name;
                                productId = quoteLineProductId[0].id.toLowerCase().replace('{', '').replace('}', '');
                                }

                            if (productId != null && businessUnit != null && warehouseNumber != null) {
                                var entity = "/products(" + productId + ")";
                                var columns = "productnumber,msdyn_fieldserviceproducttype";
                                Sdk.request("GET", entity + "?$select=" + columns)
                                    .then(function (request) {
                                        if (request != null) {
                                         //   debugger;
                                            var result = JSON.parse(request.response);//getting results 
                                                counter++;
                                                var productType = result["msdyn_fieldserviceproducttype"];
                                                if (productType == 690970000) { // Inventory : 690970000                                                    
                                                    var productNumber = result["productnumber"];
                                                    if (productNumber != null) {
                                                        productMessages.push(QuoteLine.CheckOnhandInventory(productNumber, businessUnit, warehouseNumber, productName, quantity, counter));
                                                    }
                                                    if (productMessages.length > 0) {
                                                        var productMessage = productMessages.join(".  ");

                                                        if (productMessage != '')
                                                            _formContext.ui.setFormNotification("Product inventory check : " + productMessage, "INFORMATION", "OnHandNotify");
                                                        else
                                                            _formContext.ui.clearFormNotification("OnHandNotify");
                                                    }
                                                }
                                            else
                                                _formContext.ui.clearFormNotification("OnHandNotify");
                                        }
                                        else
                                            _formContext.ui.setFormNotification("OnHand inveventory check failed for Product: " + productName + ". Product Number is blank.", "ERROR", "OnHandNotify");
                                    })
                            }

                        }
                    })
                    .catch(function (err) {
                        Xrm.Utility.alertDialog(err.message);

                    })
            }
        }

    };

    this.ShowPeriodByProjectType = function (exeContext) {
        var _formContext = exeContext.getFormContext();
        var quoteId = _formContext.getAttribute("quoteid");

        if (quoteId != null && quoteId.getValue() != null) {

            quoteId = quoteId.getValue()[0].id.replace('{', '').replace('}', '');

            if (quoteId != null) {
                var filter = "&$filter=quoteid eq " + quoteId + "";
                var query = "?$select=infy_projecttype" + filter;
                //get project type on Quote
                Sdk.request("GET", "/quotes" + query)
                    .then(function (request) {
                        if (request != null) {
                            results = JSON.parse(request.response)//getting results
                            if (results.value.length != 0) {

                                var projectType = results.value[0]["infy_projecttype"];

                                return { projectType };

                            }

                        }
                    })

                    .then(function (request) {
                        if (request != undefined) {

                            var projectType = request.projectType;
                            if (projectType != null) {
                                var revenueSchedule = _formContext.getControl("infy_revenueschedule");
                                var period = _formContext.getControl("infy_period");
                                var partAvailability = _formContext.getControl("infy_partavailability");
                                //var typeOfProduct = _formContext.getControl("infy_producttype");
                                var unitPrice = _formContext.getControl("infy_unitprice");
                                //based on type display the fields
                                if (projectType == "854360001") {//rental
                                    _formContext.getControl("priceperunit").setLabel("Rental Total Price");
                                    _formContext.getAttribute("ispriceoverridden").setValue(true);
                                    QuoteLine.LockFields();

                                    QuoteLine.SetFieldVisible(unitPrice, true);
                                    QuoteLine.SetFieldVisible(period, true);
                                    QuoteLine.SetFieldVisible(revenueSchedule, true);
                                    _formContext.getAttribute("infy_partavailability").setValue(true);
                                    QuoteLine.SetFieldVisible(partAvailability, false);
                                    //_formContext.getAttribute("infy_producttype").setValue(null);
                                    //QuoteLine.SetFieldVisible(typeOfProduct, false);
                                }
                                else
                                    if (projectType == "854360003") {//Other
                                        _formContext.getAttribute("infy_unitprice").setRequiredLevel("none");
                                        QuoteLine.SetFieldVisible(unitPrice, false);
                                        _formContext.getAttribute("infy_revenueschedule").setValue(null);
                                        QuoteLine.SetFieldVisible(revenueSchedule, false);
                                        _formContext.getAttribute("infy_period").setValue(null);
                                        QuoteLine.SetFieldVisible(period, false);
                                        QuoteLine.SetFieldVisible(partAvailability, true);
                                        // QuoteLine.SetFieldVisible(typeOfProduct, true);
                                    }
                                    else {//T&M and Fixed
                                        _formContext.getAttribute("infy_period").setValue(null);
                                        QuoteLine.SetFieldVisible(period, false);
                                        _formContext.getAttribute("infy_unitprice").setRequiredLevel("none");
                                        QuoteLine.SetFieldVisible(unitPrice, false);
                                        _formContext.getAttribute("infy_revenueschedule").setValue(null);
                                        QuoteLine.SetFieldVisible(revenueSchedule, false);
                                        _formContext.getAttribute("infy_partavailability").setValue(true);
                                        _formContext.getAttribute("infy_producttype").setValue(null);
                                        // QuoteLine.SetFieldVisible(partAvailability, false);
                                        //QuoteLine.SetFieldVisible(typeOfProduct, false);

                                    }
                            }
                        }

                    })
                    .catch(function (err) {
                        Xrm.Utility.alertDialog(err.message);

                    })
            }
        }
    };

    this.LockFields = function () {
        setTimeout(function () {
            Xrm.Page.ui.controls.get("ispriceoverridden").setDisabled(true);
            Xrm.Page.ui.controls.get("priceperunit").setDisabled(true);
        }, 3000);

    };

    this.LockFieldsOnChangeofProduct = function (exeContext) {
        var _formContext = exeContext.getFormContext();
        var quoteId = _formContext.getAttribute("quoteid");
        if (quoteId != null && quoteId.getValue() != null) {
            quoteId = quoteId.getValue()[0].id.replace('{', '').replace('}', '');
            if (quoteId != null) {
                var filter = "&$filter=quoteid eq " + quoteId + "";
                var query = "?$select=infy_projecttype" + filter;
                //get project type on Quote
                Sdk.request("GET", "/quotes" + query)
                    .then(function (request) {
                        if (request != null) {
                            results = JSON.parse(request.response)//getting results
                            if (results.value.length != 0) {
                                var projectType = results.value[0]["infy_projecttype"];
                                if (projectType == "854360001")
                                    QuoteLine.LockFields();

                            }

                        }
                    })
            }
        }

    };

    this.SetFieldVisible = function (control, visibility) {
        if (control == null || control == undefined) {
            return;
        }
        control.setVisible(visibility);
    };

    this.GetUnitPriceByProduct = function (exeContext) {

        var _formContext = exeContext.getFormContext();
        var productValue = _formContext.getAttribute("productid");
        var quoteId = _formContext.getAttribute("quoteid");

        if (quoteId != null && quoteId.getValue() != null) {

            quoteId = _formContext.getAttribute("quoteid").getValue()[0].id.replace('{', '').replace('}', '');
            if (quoteId != null) {

                var entityLogicalName = "/quotes(" + quoteId + ")";
                var columnsToRetrieve = "_pricelevelid_value";
                //get Quote price List 
                Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                    .then(function (request) {
                        if (request != null) {

                            result = JSON.parse(request.response);
                            var productPriceList = result["_pricelevelid_value"];
                            return { productPriceList };
                        }
                    })
                    .then(function (request) {
                        if (request != null) {

                            var priceList = request.productPriceList.replace('{', '').replace('}', '');
                            if (productValue.getValue() != null) {

                                productValue = productValue.getValue()[0].id.replace('{', '').replace('}', '');

                                if (priceList != null && productValue != null) {

                                    var filter = "&$filter=_productid_value eq " + productValue + " and  _pricelevelid_value eq " + priceList + "";
                                    var query = "?$select=amount" + filter;
                                    //basaed on selected Product and Quote Pricelist get amount 
                                    Sdk.request("GET", "/productpricelevels" + query)
                                        .then(function (productrequest) {
                                            if (productrequest != null) {
                                                results = JSON.parse(productrequest.response)//getting results
                                                if (results.value.length != 0) {

                                                    var amount = results.value[0]["amount"];
                                                    _formContext.getAttribute("infy_unitprice").setValue(amount);
                                                   
                                                }

                                            }
                                        })

                                }
                            }

                        }
                    })
                    .catch(function (err) {
                        Xrm.Utility.alertDialog(err.message);

                    })

            }
        }

    };

    this.GetSheduleValue = function (exeContext) {
        var _formContext = exeContext.getFormContext();
        var revenueSchedule = _formContext.getAttribute("infy_revenueschedule").getValue();
        var unitprice = _formContext.getAttribute("infy_unitprice").getValue();
        if (revenueSchedule != null) {
            revenueSchedule = revenueSchedule[0].id.replace('{', '').replace('}', '');
            if (revenueSchedule != null) {
                var entityLogicalName = "/infy_revenueschedules(" + revenueSchedule + ")";
                var columnsToRetrieve = "infy_recurrencevalue";
                Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                    .then(function (request) {
                        if (request != null) {
                            result = JSON.parse(request.response);
                            var value = result["infy_recurrencevalue"];
                            if (value != null) {
                                _formContext.getAttribute("infy_period").setValue(value);
                                var period = _formContext.getAttribute("infy_period").getValue();
                                if (unitprice != null && period != null)
                                    _formContext.getAttribute("priceperunit").setValue(unitprice * period);
                            }
                        }
                    })
                    .catch(function (err) {
                        Xrm.Utility.alertDialog(err.message);

                    })
            }
            else
                _formContext.getAttribute("infy_period").setValue(null);
        }
        else
            _formContext.getAttribute("infy_period").setValue(null);
    };

     this.CalculateTotalPrice = function (exeContext) {//based on unitprice and period calculate price field value
 
        var _formContext = exeContext.getFormContext();
        var unitprice = _formContext.getAttribute("infy_unitprice").getValue();
        var period = _formContext.getAttribute("infy_period").getValue();
        if (unitprice !== null && period !== null)
            _formContext.getAttribute("priceperunit").setValue(unitprice * period);
        else {
            if (unitprice !== null) {
                _formContext.getAttribute("priceperunit").setValue(unitprice);
            }
        }
    };
//On change of product calculate the rental total price 
this.SetPriceValue = function (exeContext) {
_formContext = exeContext.getFormContext();
        var productId = _formContext.getAttribute("productid");
        if (productId != null) {
            productId = productId.getValue();
            if (productId != null)
            setTimeout(function () { QuoteLine.CalculateTotalPrice(exeContext);QuoteLine.CalculateTax(exeContext);}, 3000);
            
        }
    };

    //Filter Parent Product lookup to only include quote line products except the current one
    this.FilterParentProducts = function (exeContext) {
        try {
            var _formContext = exeContext.getFormContext();
            var productIds = [];
            var productId;
            var filterXML;

            var productValue = _formContext.getAttribute("productid").getValue();
            if (productValue)
                productId = productValue[0].id.replace('{', '').replace('}', '');
            var quoteValue = _formContext.getAttribute("quoteid").getValue();
            if (quoteValue != null) {
                var quoteId = quoteValue[0].id.replace('{', '').replace('}', '');
                if (quoteId != null) {
                    var filter = "&$filter=_quoteid_value eq " + quoteId + "";
                    var query = "?$select=_productid_value" + filter;

                    Sdk.request("GET", "/quotedetails" + query)
                        .then(function (request) {
                            if (request != null) {
                                results = JSON.parse(request.response)//getting results                            
                                for (var i = 0; i < results.value.length; i++) {
                                    var quoteProduct = results.value[i]["_productid_value"];
                                    if (quoteProduct != null) {
                                        if (productId != null) {
                                            if (productId.toLowerCase() != quoteProduct.toLowerCase())
                                                productIds.push(quoteProduct.toLowerCase());
                                        }
                                        else
                                            productIds.push(quoteProduct.toLowerCase());
                                    }
                                }
                                if (productIds.length >= 0) {
                                    var ids = "<value>00000000-0000-0000-0000-000000000000</value>";
                                    for (var i = 0; i < productIds.length; i++) {
                                        if (i == 0) {
                                            ids = "";
                                        }
                                        ids += "<value>" + productIds[i] + "</value>";
                                    }

                                    filterXML = ("<filter type='and'>\
                                        <condition attribute='productid' operator='in'>" +
                                        ids +
                                        "  </condition>\
                                        </filter>");
                                    if (filterXML != null) {
                                        _formContext.getControl("infy_parentproduct").addPreSearch(
                                            function () {
                                                _formContext.getControl("infy_parentproduct").addCustomFilter(filterXML);
                                            })
                                    }

                                }

                            }

                        })
                        .catch(function (err) {
                            Xrm.Utility.alertDialog(err.message);

                        });
                }
            }

        } catch (e) {

            throw new Error("FilterParentProducts" + e.message);

        }
    };

    // Calcualte Tax On change of - Exiting product, Price Per Unit and Quantity.
    this.CalculateTax = function (exeContext) {
debugger;
        var _formContext = exeContext.getFormContext();
        var taxable = _formContext.getAttribute("msdyn_taxable");
        if (taxable.getValue() != null) {
            // Check if Taxable is true then only calculate tax else do not calcualte tax
            if (taxable.getValue() == 1) // true
            {
                var quoteId = _formContext.getAttribute("quoteid");
                if (quoteId != null && quoteId.getValue() != null) {
                    quoteId = _formContext.getAttribute("quoteid").getValue()[0].id.replace('{', '').replace('}', '');
                    if (quoteId != null) {
                        var entityLogicalName = "/quotes(" + quoteId + ")";
                        var columnsToRetrieve = "_infy_taxcode_value";
                        // Get Tax Code from Opportunity
                        Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                            .then(function (request) {
                                if (request != null) {
                                    result = JSON.parse(request.response);
                                    var taxCodeLookup = result["_infy_taxcode_value"];
                                    return { taxCodeLookup };
                                }
                            })
                            .then(function (request) {
                                if (request != null) {
                                    // Check if Tax Code is selected on Header or Not. 
                                    if (request.taxCodeLookup != null) {
                                        var taxCodeId = request.taxCodeLookup.replace('{', '').replace('}', '');
                                        var entityLogicalName = "/msdyn_taxcodes(" + taxCodeId + ")";
                                        var columnsToRetrieve = "msdyn_taxrate";
                                        // Get Tax Rate from Tax Code
                                        Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                                            .then(function (taxCodeRequest) {
                                                if (taxCodeRequest != null) {
                                                    results = JSON.parse(taxCodeRequest.response)
                                                    if (results != null) {
                                                        debugger;
                                                        var taxRate = results["msdyn_taxrate"];
                                                        var quantity = _formContext.getAttribute("quantity").getValue();
                                                        var pricePerUnit = _formContext.getAttribute("priceperunit").getValue();
                                                        if (quantity == null)
                                                            quantity = 0; // Set Default Quantity
                                                        if (pricePerUnit == null)
                                                            pricePerUnit = 0; // Set Default Price Per Unit
                                                        var tax = quantity * pricePerUnit * (taxRate / 100); // Calcualte Tax
                                                        _formContext.getAttribute("tax").setValue(tax); // Set Tax value
                                                    }
                                                    else {
                                                        _formContext.getAttribute("tax").setValue(0);
                                                    }
                                                }
                                            })
                                    }
                                    else {
                                        // if Tax Code is not selected on header then set tax to 0
                                        _formContext.getAttribute("tax").setValue(0);
                                    }
                                }
                            })
                            .catch(function (err) {
                                Xrm.Utility.alertDialog(err.message);
                            })
                    }
                }
            }
        }
    };

    this.SetTaxValue = function (exeContext) {
debugger;
        setTimeout(function () { QuoteLine.CalculateTax(exeContext); }, 2000);
    };

 this.CheckStatus = function (exeContext) {
     
        var _formContext = exeContext.getFormContext();
        var quoteId = _formContext.getAttribute("quoteid");

        if (quoteId != null && quoteId.getValue() != null) {
            quoteId = _formContext.getAttribute("quoteid").getValue()[0].id.replace('{', '').replace('}', '');
            var entityLogicalName = "/quotes(" + quoteId + ")";
            var columnsToRetrieve = "statecode";
            Sdk.request("GET", entityLogicalName + "?$select=" + columnsToRetrieve + "")
                .then(function (request) {
                    if (request != null) {
                        result = JSON.parse(request.response);
                        var stateCode = result["statecode"];//get schedule value
                        if (stateCode == 1 || stateCode == 2 || stateCode == 3) {//  Active/WOn/Closed quote
                            QuoteLine.ReadOnly(exeContext);
                        }
                    }
                })
                .catch(function (err) {
                    Xrm.Utility.alertDialog(err.message);

                })           

        }

    }

 this.ReadOnly = function (exeContext) {
        var _formContext = exeContext.getFormContext()
        setTimeout(function () {            
            _formContext.ui.controls.forEach(function (control, i) {
                if (!control.getDisabled()) {
                    control.setDisabled(true);
                }
            });
        }, 2000);

    };
}).call(QuoteLine);

