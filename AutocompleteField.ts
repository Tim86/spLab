module Tahoe.Forms {
    export class AutocompleteField {
        private formFieldName: string;

        /**
         * 
         * @param autocompleteFieldName The form field that will get autocomplete
         */
        constructor(autocompleteFieldName) {
            this.formFieldName = autocompleteFieldName;
        }

        /**
         * 
         * @param sourceListName
         * @param sourceFieldName
         */
        bindAutocomplete(sourceListName, sourceFieldName) {
            var self = this;
            var input = $('input[id^="' + self.formFieldName + '"]');
            var clientContext = SP.ClientContext.get_current();
            var list = clientContext.get_web().get_lists().getByTitle(sourceListName);

            // here we use jquery ui autocomplete function
            input.autocomplete({
                // the important parameter, here we set the source as a function
                source(request, response) {
                    var camlQuery = new SP.CamlQuery();
                    camlQuery.set_viewXml('' +
                        '<View><Query><Where>' +
                        '<Contains>' +
                        '  <FieldRef Name="' + sourceFieldName + '"/>' +
                        '  <Value Type="Text">' + request.term + '</Value>' +
                        '</Contains>' +
                        '</Where></Query></View>');
                    var items = list.getItems(camlQuery);
                    clientContext.load(items);
                    // here we query sharepoint async for our list data
                    clientContext.executeQueryAsync(() => {
                        console.log("match count: " + items.get_count());
                        var resultData = [];
                        var productListEnumerator = items.getEnumerator();
                        while (productListEnumerator.moveNext()) {
                            var listItem = productListEnumerator.get_current();
                            var sourceFieldValue = listItem.get_item(sourceFieldName);
                            resultData.push(sourceFieldValue);
                        }
                        // here we tell autocomplete that where are finished and have a result array
                        response(resultData);
                    }, (sender, args) => {
                        console.log("error cougth:" + args.get_message());
                    });
                },
                minLength: 2
            });

        }
    }
}

var autocompleteField = new OneMed.Customers.AutocompleteField("Customers", "OM_ComplaintCustomerName");
autocompleteField.bindAutocomplete("OM_CustomerName", "OM_ComplaintCustomerID", "OM_CustomerID"); // OM_ComplaintCustomerID is not on the companies list!?

//SP.SOD.executeOrDelayUntilScriptLoaded(() => {
//    // or wait untill csr has done its magic
//    $(document).ready(() => {
//        var autocompleteField = new OneMed.Customers.AutocompleteField("Customers", "OM_ComplaintCustomerName");
//        autocompleteField.bindAutocomplete("OM_CustomerName", "OM_ComplaintCustomerID", "OM_CustomerID"); // OM_ComplaintCustomerID is not on the companies list!?
//    });
//}, "tahoe.init.js");
