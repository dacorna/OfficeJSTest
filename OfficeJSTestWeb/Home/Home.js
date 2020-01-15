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

    function updateStocks() {
        // can reference host-specific api 'Excel' directly as opposed to 'Office'
        // run anonymous function that accepts a context parameter 
        Excel.run(function (ctx) {
            // get range object, must have a named range called Stocks
            var range = ctx.workbook.names.getItem("Stocks").getRange();
            var rangeValues = range.getRow(1).getBoundingRect(range.getLastCell());
            var stockNamesRange = range.getColumn(0);
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
                app.showNotification("Values read", range.values);
            })
        }).catch(function (error) {
            app.showNotification("Error", error);
        })
    }
})();