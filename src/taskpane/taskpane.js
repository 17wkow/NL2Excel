/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Global variable to store the latest generated formula
let lastGeneratedFormula = "";

// Import Gemini SDK
import { GoogleGenerativeAI } from "@google/generative-ai";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("generate-btn").onclick = generateFormula;
        document.getElementById("insert-btn").onclick = insertFormula;
        document.getElementById("save-key-btn").onclick = saveApiKey;

        // Load saved key if available, or use env var
        const savedKey = localStorage.getItem("gemini_api_key") || process.env.GEMINI_API_KEY;
        if (savedKey) {
            document.getElementById("api-key").value = savedKey;
        }
    }
});

async function generateFormula() {
    const description = document.getElementById("description").value;
    const apiKey = document.getElementById("api-key").value;
    const messageArea = document.getElementById("message-area");
    const resultSection = document.getElementById("result-section");
    const formulaPreview = document.getElementById("formula-preview");

    if (!description) {
        messageArea.innerText = "Please enter a description.";
        return;
    }
    if (!apiKey) {
        messageArea.innerText = "Please enter your Gemini API Key in settings.";
        return;
    }

    messageArea.innerText = "Generating...";
    resultSection.classList.add("hidden");

    try {
        const formula = await callLLM(description, apiKey);
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
    const key = document.getElementById("api-key").value;
    localStorage.setItem("gemini_api_key", key);
    document.getElementById("message-area").innerText = "API Key saved.";
    setTimeout(() => {
        document.getElementById("message-area").innerText = "";
    }, 2000);
}

async function callLLM(prompt, apiKey) {
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });

    const systemPrompt = "You are an expert Excel assistant. User will ask for a formula. Output ONLY the Excel formula, starting with =. Do not explain. If the request is not clear, guess the most likely standard formula.";

    const result = await model.generateContent([systemPrompt, prompt]);
    const response = await result.response;
    let content = response.text().trim();

    // Clean up if it wrapped in markdown code blocks
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

