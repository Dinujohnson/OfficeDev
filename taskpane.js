// taskpane.js

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // Add an event handler to the custom button
        OfficeExtension.ExtensionHelpers.registerOnMessageHandler(function (args) {
            if (args.message === "addWatermark") {
                addWatermark();
            }
        });
    }
});

async function addWatermark() {
    try {
        // Set coercion type to text since we want to apply formatting
        const options = { coercionType: Office.CoercionType.Text };

        // Construct the watermark text with formatting
        const watermarkText = `
            <div style="font-size: 120px; font-weight: bold; color: #00FF00;">CONFIDENTIAL</div>
        `;

        // Insert the watermark HTML with styling into the current selection
        await Office.context.document.setSelectedDataAsync(watermarkText, options);
    } catch (error) {
        console.error("Error inserting watermark:", error);
    }
}
