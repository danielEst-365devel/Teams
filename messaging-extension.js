// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

// Function to handle message extension query
async function handleMessagingExtensionQuery(query) {
    // Fetch translation results using the Microsoft Translator API (replace YOUR_TRANSLATOR_API_KEY)
    const translatorApiKey = 'YOUR_TRANSLATOR_API_KEY';
    const userQuery = query.message.text;

    try {
        const translationResponse = await fetch(`https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to=en`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Ocp-Apim-Subscription-Key': translatorApiKey,
            },
            body: JSON.stringify([{ text: userQuery }]),
        });

        const translationData = await translationResponse.json();

        const translationResult = translationData[0].translations[0].text;

        // Return the translation as a messaging extension result
        const result = {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [
                {
                    content: {
                        title: `Translation to English:`,
                        text: translationResult,
                    },
                },
            ],
        };

        microsoftTeams.tasks.submitTaskResults([result]);
    } catch (error) {
        console.error('Error translating message:', error);
    }
}

// Register messaging extension query handler
microsoftTeams.tasks.registerOnQuery(context => {
    handleMessagingExtensionQuery(context);
});
