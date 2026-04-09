/// <reference path="./custom-annotator.js" />
/// <reference path="./types.js" />

/**
 * @param {localStorageCustomAssistantItem} assistantData
 * @constructor
 * @extends CustomAnnotator
 */
function AssistantReplace(annotationPopup, assistantData) {
    CustomAnnotator.call(this, annotationPopup, assistantData);
}
AssistantReplace.prototype = Object.create(CustomAnnotator.prototype);
AssistantReplace.prototype.constructor = AssistantReplace;

Object.assign(AssistantReplace.prototype, {
    /**
     * @param {string} text
     * @param {ReplaceAiResponse[]} matches
     */
    _convertToRanges: function (paraId, text, matches) {
        const _t = this;
        let rangeId = 1;
        const ranges = [];
        for (const {
            origin,
            suggestion,
            paragraph,
            occurrence,
            confidence,
        } of matches) {
            if (origin === suggestion || confidence <= 0.7) continue;

            let count = 0;
            let searchStart = 0;

            while (searchStart < text.length) {
                const index = _t.simpleGraphemeIndexOf(
                    text,
                    origin,
                    searchStart,
                );
                if (index === -1) break;

                count++;
                if (count === occurrence) {
                    ranges.push({
                        start: index,
                        length: [...origin].length,
                        id: rangeId,
                    });
                    _t.paragraphs[paraId][rangeId] = {
                        original: origin,
                        suggestion: suggestion,
                    };
                    ++rangeId;
                    break;
                }
                searchStart = index + 1;
            }
        }
        return ranges;
    },

    /**
     * @param {string} text
     * @returns {string}
     */
    _createPrompt: function (text) {
        let prompt = `You are a multi-disciplinary text analysis and transformation assistant.
	  Your task is to analyze text based on user's specific criteria and provide intelligent corrections.
	
	  MANDATORY RULES:
		1. UNDERSTAND the user's intent from their criteria.
		2. FIND all words matching the criteria.
		3. For EACH match you find:
		  - Provide the exact quote.
		  - SUGGEST appropriate replacements.
		  - Explain WHY it matches the criteria.
		  - Provide position information (paragraph number).
		4. If no matches are found, return an empty array: [].
		5. Format your response STRICTLY in JSON format.
		6. Support multiple languages (English, Russian, etc.)

	  Response format - return ONLY this JSON array with no additional text:
		[
		  {
			"origin": "exact text fragment that matches the query",
      		"suggestion": "suggested replacement (plain text)",
			"paragraph": paragraph_number,
			"occurrence": 1,
			"confidence": 0.95
		  }
		]

	  Guidelines for each field:
		- "origin": EXACT UNCHANGED original text fragment. Do not fix anything in this field.
		- "suggestion": Your suggested replacement for the fragment.
			* Ensure it aligns with the user's criteria.
			* Maintain coherence with surrounding text.
		- "paragraph": Paragraph number where the fragment is found (0-based index)
		- "occurrence": Which occurrence of this sentence if it appears multiple times (1 for first, 2 for second, etc.)
		- "confidence": Value between 0 and 1 indicating certainty (1.0 = completely certain, 0.5 = uncertain)
	  
	  CRITICAL:
		- Output should be in the exact this format
		- No any comments are allowed

	  CRITICAL - Output Format:
		- Return ONLY the raw JSON array, nothing else
		- DO NOT wrap the response in markdown code blocks (no \`\`\`json or \`\`\`)
		- DO NOT include any explanatory text before or after the JSON
		- DO NOT use escaped newlines (\\n) - return the JSON on a single line if possible
		- The response should start with [ and end with ]
	  `;
        prompt +=
            "\n\nUSER REQUEST:\n```" + this.assistantData.query + "\n```\n\n";

        prompt += "TEXT TO ANALYZE:\n```\n" + text + "\n```\n\n";

        prompt += `Please analyze this text and find all fragments that match the user's request. Be thorough but precise.`;

        return prompt;
    },

    /**
     * @param {string} paraId
     * @param {string} rangeId
     * @returns {ReplaceInfoForPopup}
     */
    getInfoForPopup: function (paraId, rangeId) {
        let _s = this.getAnnotation(paraId, rangeId);
        let suggested = _s["suggestion"];
        if (suggested.indexOf('</strong>') === -1) {
            suggested = `<strong>${suggested}</strong>`;
        }
        return {
            original: _s["original"],
            suggested: suggested,
            type: this.type,
        };
    },

    /**
     * @param {string} paraId
     * @param {string} rangeId
     */
    onAccept: async function (paraId, rangeId) {
        await CustomAnnotator.prototype.onAccept.call(this);
        let text = this.getAnnotation(paraId, rangeId)["suggestion"];

        await Asc.Editor.callMethod("StartAction", ["GroupActions"]);

        let range = this.getAnnotationRangeObj(paraId, rangeId);
        await Asc.Editor.callMethod("SelectAnnotationRange", [range]);

        Asc.scope.text = text;
        await Asc.Editor.callCommand(function () {
            Api.ReplaceTextSmart([Asc.scope.text]);
            Api.GetDocument().RemoveSelection();
        });

        await Asc.Editor.callMethod("RemoveAnnotationRange", [range]);
        await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
        await Asc.Editor.callMethod("FocusEditor");
    },
});
