/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Global variable to store the latest generated formula
let lastGeneratedFormula = "";

// Import Gemini SDK
// import { GoogleGenerativeAI } from "@google/generative-ai";

const PROVIDERS = {
    gemini: { name: "Gemini", keyStorage: "gemini_api_key" },
    openai: { name: "OpenAI", keyStorage: "openai_api_key" }
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("generate-btn").onclick = generateFormula;
        document.getElementById("insert-btn").onclick = insertFormula;
        document.getElementById("save-key-btn").onclick = saveApiKey;
        document.getElementById("provider-select").onchange = updateProviderUI;

        // Initialize UI
        updateProviderUI();
    }
});

function getSelectedProvider() {
    return document.getElementById("provider-select").value;
}

function updateProviderUI() {
    const provider = getSelectedProvider();
    const config = PROVIDERS[provider];

    document.getElementById("api-key-label").innerText = `${config.name} API Key:`;

    // Load saved key
    const savedKey = localStorage.getItem(config.keyStorage) || "";
    document.getElementById("api-key").value = savedKey;
}

async function generateFormula() {
    const description = document.getElementById("description").value;
    const apiKey = document.getElementById("api-key").value;
    const provider = getSelectedProvider();
    const messageArea = document.getElementById("message-area");
    const resultSection = document.getElementById("result-section");
    const formulaPreview = document.getElementById("formula-preview");

    if (!description) {
        messageArea.innerText = "Please enter a description.";
        return;
    }
    if (!apiKey) {
        messageArea.innerText = `Please enter your ${PROVIDERS[provider].name} API Key in settings.`;
        return;
    }

    messageArea.innerText = "Generating...";
    resultSection.classList.add("hidden");

    try {
        const formula = await callLLM(description, apiKey, provider);
        lastGeneratedFormula = formula;
        formulaPreview.innerText = formula;
        resultSection.classList.remove("hidden");
        messageArea.innerText = "";
    } catch (error) {
        console.error(error);
        messageArea.innerText = "Error: " + error.message;
    }
}

async function insertFormula() {
    if (!lastGeneratedFormula) return;

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.formulas = [[lastGeneratedFormula]];
            await context.sync();
        });
    } catch (error) {
        // Fallback for older Excel versions or other contexts
        Office.context.document.setSelectedDataAsync(
            lastGeneratedFormula,
            { coercionType: Office.CoercionType.Text },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    document.getElementById("message-area").innerText = "Error inserting: " + asyncResult.error.message;
                }
            }
        );
    }
}

function saveApiKey() {
    const provider = getSelectedProvider();
    const key = document.getElementById("api-key").value;
    localStorage.setItem(PROVIDERS[provider].keyStorage, key);

    document.getElementById("message-area").innerText = `${PROVIDERS[provider].name} Key saved.`;
    setTimeout(() => {
        document.getElementById("message-area").innerText = "";
    }, 2000);
}

async function callLLM(prompt, apiKey, provider) {
    const systemPrompt = "You are an expert Excel assistant. User will ask for a formula. Output ONLY the Excel formula, starting with =. Do not explain. If the request is not clear, guess the most likely standard formula.";

    if (provider === "gemini") {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{
                    parts: [{ text: `${systemPrompt}\n\n${prompt}` }]
                }]
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error((errorData.error && errorData.error.message) || `Gemini Error: ${response.status}`);
        }

        const data = await response.json();
        const text = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0] && data.candidates[0].content.parts[0].text) || "";
        return cleanFormula(text);
    } else if (provider === "openai") {
        const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "gpt-4o-mini",
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: prompt }
                ]
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            const msg = (errorData.error && errorData.error.message) || `OpenAI Error: ${response.status}`;
            throw new Error(msg);
        }

        const data = await response.json();
        return cleanFormula(data.choices[0].message.content);
    }
}

function cleanFormula(text) {
    let content = text.trim();
    if (content.startsWith("```")) {
        content = content.replace(/^```(excel)?\s*/, "").replace(/\s*```$/, "");
    }
    return content;
}


// DEBUG: List available models
async function listAvailableModels() {
    const apiKey = document.getElementById("api-key").value;
    if (!apiKey) {
        console.log("No API key found for debug listing.");
        return;
    }

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        const data = await response.json();
        console.log("DEBUG: Available Models for this key:", data);
        if (data.models) {
            const names = data.models.map(m => m.name);
            console.log("DEBUG: Model Names:", names);
            document.getElementById("message-area").innerText += "\n(Check console for available models)";
        }
    } catch (e) {
        console.error("DEBUG: Failed to list models", e);
    }
}

// Call on saved key
document.getElementById("save-key-btn").addEventListener('click', listAvailableModels);
// Try on load
setTimeout(listAvailableModels, 2000); // Give time for value to populate

