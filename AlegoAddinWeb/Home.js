
(function () {
    "use strict";

    var messageBanner;

    // A função inicializar deverá ser executada cada vez que uma nova página for carregada.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // inicialize o mecanismo de notificação do FabricUI e oculte-o
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // Se não estiver usando o Word 2016, use a lógica de fallback.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("Este add-in formata o documento para os padrões da Alego.");
                $('#button-text').text("Exibir!");
                $('#button-desc').text("Exibir o texto selecionado");
                
                $('#header-button').click(createHeader);
                return;
            }

            $("#template-description").text("Este add-in formata o documento para os padrões da Alego");
            $('#button-text').text("Criar cabeçalho!");
            $('#button-desc').text("Insere o cabeçalho no padrão Alego.");
            

            // Adicione um manipulador de eventos de clique ao botão.
            $('#header-button').click(createHeader);
        });
    };

    function createHeader() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections;

            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');

            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Create a proxy object the primary header of the first section. 
                // Note that the header is a body object.
                var myHeader = mySections.items[0].getHeader("primary");

                // Queue a command to insert text at the end of the header.
                // myHeader.insertText("This is a header.", Word.InsertLocation.end);

                var conte = myHeader.getHtml();

                // Queue a command to wrap the header in a content control.
                myHeader.insertContentControl();

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log(conte);
                });
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Erro:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Função auxiliar para exibir notificações
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
