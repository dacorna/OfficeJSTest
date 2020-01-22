'use strict';

(function () {

    // onReady is recommended over initialize, can call this in different places with different callbacks 
    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            app.initialize();

            $('#update-stocks').click(updateStocks);
            //$('#set-color').click(setColor);
        });
    });

    function setColor() {
        Excel.run(function (context) {
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = 'green';

            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function initDatePicker() {
        var $date = $('#date');
        $date.datepicker();
        $date.datepicker("option", "dateFormat", "yy-mm-dd");

        $date.change(function () {
            Office.context.document.settings.set("Date", $date.val());
            Office.context.document.settings.saveAsync();
        });

        $date.datepicker("setDate", Office.context.document.settings.get("Date"));
    }

    function updateStocks() {
        // can reference host-specific api 'Excel' directly as opposed to 'Office'
        // run anonymous function that accepts a context parameter 
        Excel.run(function (ctx) {
            // get range object, must have a named range called Stocks
            var range = ctx.workbook.names.getItem("Stocks").getRange();
            // if we wanted to access range.values later we have to load "values" first 
            //range.load("values");

            var rangeValues = range.getRow(1).getBoundingRect(range.getLastCell());

            rangeValues.format.fill.clear();
            rangeValues.getColumn(1).getBoundingRect(rangeValues.getLastColumn()).clear(Excel.ClearApplyTo.contents);

            var stockNamesRange = rangeValues.getColumn(0);
            stockNamesRange.load("values");

            // these two things are 100% identical
            //range.load("values");
            //ctx.load(range, "values");
            // ------------------------------------

            // this context represents the command queue you'll be executing 
            // chain of commands sent back to document and state gets synchronized 
            // ctx.sync() returns a promise object 
            // when the sync is complete, I then want to go and do something 
            return ctx.sync().then(function () {               
                app.showNotification("Values read", stockNamesRange.values);

                var stocks = stockNamesRange.values.map(function (item) {
                    return '"' + item[0] + '"';
                });
                var url = '//query.yahooapis.com/v1/public/yql';
                var data = encodeURIComponent('select * from yahoo.finance.historicaldata where symbol in (' + stocks.join(',') + ') and startDate = "' + $('#date').val() + '" and endDate = "' + $('#date').val() + '"');

                return Q.Promise(function (resolve, reject) {
                    $.getJSON(url, 'q=' + data + "&env=http%3A%2F%2Fdatatables.org%2Falltables.env&format=json")
                        .done(function (result) {
                            resolve(result);
                        })
                        .fail(function (error) {
                            reject(error.statusText);
                        });
                });              
            })
                .then(function (result) {
                    var stockValues;
                    try {
                        stockValues = result.query.results.quote;
                        if (stockValues.length == 0) {
                            throw new Error();
                        }
                    } catch (e) {
                        throw new Error("Could not retrieve stock values from the server for the specified day");
                    }

                    var dataArray = stockValues.map(function (item, index) {
                        var row = rangeValues.getRow(index);
                        if (item.Close < item.Open) {
                            row.format.fill.color = "red";
                        } else if (item.Close > item.Open) {
                            row.format.fill.color = "green";
                        }
                        return [item.Symbol, item.Open, item.Close];
                    })

                    var rangeToWriteTo = rangeValues.getRow(0).getBoundingRect(
                        rangeValues.getRow(stockValues.length - 1));
                    rangeToWriteTo.values = dataArray;
                    return ctx.sync();
                })
        }).catch(function (error) {
            app.showNotification("Error", error);
        })
    }
})();