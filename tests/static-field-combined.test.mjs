import assert from "node:assert/strict";
import test from "node:test";

import { buildStaticAnswers, matchStaticFieldKey } from "../server/profiles/static-fields.js";

test("combined preferred first-and-last-name fields receive the complete saved name", () => {
  const field = {
    id: "preferred_name",
    label: "What is your preferred first and last name?",
    name: "preferred_name",
    autocomplete: "",
    placeholder: "",
    type: "text",
    options: []
  };

  assert.equal(matchStaticFieldKey(field), "full_name");
  assert.equal(buildStaticAnswers([field], {
    first_name: "Anton",
    last_name: "Tjakan",
    full_name: "Anton Tjakan"
  }).get(field.id)?.value, "Anton Tjakan");
});

test("combined city-state-country prompts receive the complete saved location", () => {
  const field = {
    id: "location",
    label: "What location are you based in? (city, state, and country)",
    name: "location",
    autocomplete: "",
    placeholder: "",
    type: "text",
    options: []
  };

  assert.equal(matchStaticFieldKey(field), "location");
  assert.equal(buildStaticAnswers([field], {
    city: "Warsaw",
    country: "Poland",
    location: "Warsaw, Poland"
  }).get(field.id)?.value, "Warsaw, Poland");
});
