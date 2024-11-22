// Version 1.0.2 - Add processor.js to loading sequence


async function loadAllScripts() {
    try {
        await loadJs("./js/ZohoEmbededAppSDK.min.js")
        await loadJs('./js/common.js');
        await loadJs('./js/zohoapi.js');
        await loadJs('./js/varEngine.js');
        await loadJs('./js/generator.js');
        await loadJs('./js/processor.js');
        await loadJs('./js/main.js');
        console.log('All scripts loaded successfully');
    } catch (error) {
        console.error('Error loading scripts:', error);
    }
}


(async () => {
    await loadAllScripts()
})();
