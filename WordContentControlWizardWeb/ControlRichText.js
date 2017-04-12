// Script for insert content control
(function () {
    Office.initialize = function (reason) {
        //add initialize function              
    };
})();

//function for insertContentControl
function insertContentControl(event) {

    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Queue a command to get the current selection and then
        // create a proxy range object with the results.
        var range = context.document.getSelection();

        // Queue a commmand to insert a content control around the selected text,
        // and create a proxy content control object. We'll update the properties
        // on the content control.
        var myContentControl = range.insertContentControl();
        
        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped a content control around the selected text.');
        });
    })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    event.completed();
}