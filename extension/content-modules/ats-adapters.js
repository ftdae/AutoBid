(() => {
  const COMMON_CONTROL_SELECTORS = [
    "[contenteditable='true'][role='textbox']",
    "[role='button'][aria-haspopup='listbox']",
    "[role='button'][aria-haspopup='menu'][aria-expanded]",
    "[data-headlessui-state][role='combobox']",
    "[data-testid*='select' i][aria-expanded]",
    "[data-testid*='dropdown' i][aria-expanded]",
    ".select2-selection",
    ".select2-choice",
    ".chosen-single"
  ];

  const COMMON_OPTION_SELECTORS = [
    "[role='option']",
    "[role='menuitemradio']",
    "[role='menuitem']",
    "[data-value]",
    "[data-radix-collection-item]",
    "[cmdk-item]",
    ".select2-results__option",
    ".select2-result-label",
    ".chosen-results li.active-result"
  ];

  const DEFINITIONS = [
    adapter("ashby", "Ashby", [/(^|\.)jobs\.ashbyhq\.com$/i], {
      markers: [".ashby-application-form-question-title", "[class*='ashby-application-form']"],
      controls: [".ashby-application-form input", ".ashby-application-form textarea", ".ashby-application-form [role='combobox']"],
      containers: ["fieldset", "[class*='_fieldEntry_']", "[class*='_container_']"],
      labels: [".ashby-application-form-question-title", "[class*='_label_']"],
      options: ["[class*='_option_'] [role='option']", "[class*='_option_'] label"]
    }),
    adapter("bamboohr", "BambooHR", [/(^|\.)bamboohr\.com$/i], {
      markers: [".fab-FormRow", "[data-bi-id*='careers']"],
      controls: [
        ".fab-FormRow input",
        ".fab-FormRow textarea",
        ".fab-FormRow select",
        ".fab-FormRow [role='combobox']",
        ".fab-FormRow [aria-haspopup='listbox']",
        ".fab-FormRow .fab-SelectToggle",
        ".fab-FormRow [data-fabric-component='SelectToggle']",
        ".fab-FormRow [class*='Select__control']",
        ".fab-Select__control",
        ".fab-SelectToggle",
        "[data-fabric-component='SelectToggle']"
      ],
      containers: [".fab-FormRow", ".fab-FormSection", "fieldset"],
      labels: [".fab-FormRow__label", ".fab-Label", "legend"],
      options: [
        ".fab-Select__option",
        "[class*='fab-Select'] [role='option']",
        "[class*='Select__option']",
        ".fab-MenuOption",
        "[data-fabric-component='MenuOption']"
      ],
      uploads: [".fab-FormRow input[type='file']", "input[type='file']"]
    }),
    adapter("gem", "Gem", [/(^|\.)gem\.com$/i, /(^|\.)gem\.co$/i], {
      markers: ["[data-testid*='application-form']", "[class*='ApplicationForm']"],
      containers: ["[data-testid*='form-field']", "[class*='FormField']", "fieldset"],
      labels: ["[data-testid*='label']", "[class*='FormLabel']"],
      options: ["[data-testid*='option']"]
    }),
    adapter("gohire", "GoHire", [/(^|\.)gohire\.io$/i], {
      markers: ["[class*='application-form']", "[data-testid*='application']"],
      containers: [".form-group", "[class*='form-field']", "fieldset"],
      labels: [".control-label", "[class*='field-label']", "legend"],
      options: ["[class*='option']", "[role='option']"]
    }),
    adapter("greenhouse", "Greenhouse", [/(^|\.)greenhouse\.io$/i, /(^|\.)greenhouse\.com$/i], {
      markers: ["#application-form", "[data-source='greenhouse']", ".application--questions"],
      controls: ["#application-form input", "#application-form textarea", "#application-form select", "#application-form [role='combobox']", ".application--questions [aria-haspopup='listbox']"],
      containers: [".field", ".application-question", "[data-testid*='field']", "fieldset"],
      labels: [".field-label", ".application-question label", "[data-testid*='label']", "legend"],
      options: ["[id*='react-select'][id*='option']", ".select__option", "[data-testid*='option']"],
      uploads: ["#resume", "input[name*='resume' i]", "input[id*='resume' i]"]
    }),
    adapter("newrocket", "NewRocket Greenhouse", [/(^|\.)newrocket\.com$/i], {
      markers: ["#application_form", "#application-form", ".application--questions", ".select2-container"],
      controls: [
        "#application_form input",
        "#application_form textarea",
        "#application_form select",
        "#application_form [role='combobox']",
        "#application_form .select2-selection",
        "#application_form .select2-choice",
        "#application_form .chosen-single",
        "#application-form input",
        "#application-form textarea",
        "#application-form select",
        "#application-form [role='combobox']",
        "#application-form .select2-selection",
        "#application-form .select2-choice",
        "#application-form .chosen-single"
      ],
      containers: [".field", ".application-question", ".field-entry", ".form-field", "fieldset"],
      labels: [".field-label", ".application-question label", ".field-entry label", "legend"],
      options: [
        ".select2-results__option",
        ".select2-result-label",
        ".chosen-results li.active-result",
        "[id*='react-select'][id*='option']",
        ".select__option",
        "[role='option']"
      ],
      uploads: ["#resume", "input[name*='resume' i]", "input[id*='resume' i]"]
    }),
    adapter("hibob", "HiBob", [/(^|\.)hibob\.com$/i, /(^|\.)bob\.co$/i], {
      markers: ["[data-testid*='candidate']", "[class*='applicationForm']"],
      containers: ["[data-testid*='field']", "[class*='fieldWrapper']", "fieldset"],
      labels: ["[data-testid*='label']", "[class*='fieldLabel']"],
      options: ["[data-testid*='option']", "[role='option']"]
    }),
    adapter("icims", "iCIMS", [/(^|\.)icims\.com$/i], {
      markers: [".iCIMS_MainWrapper", "[class*='iCIMS']"],
      controls: [".iCIMS_MainWrapper input", ".iCIMS_MainWrapper textarea", ".iCIMS_MainWrapper select", ".iCIMS_MainWrapper [role='combobox']"],
      containers: [".iCIMS_FormField", ".form-group", "fieldset"],
      labels: [".iCIMS_Label", ".control-label", "legend"],
      options: [".select2-results__option", "[role='option']"]
    }),
    adapter("jazzhr", "JazzHR", [/(^|\.)applytojob\.com$/i, /(^|\.)jazz\.co$/i], {
      markers: ["#application-form", ".apply-form"],
      controls: ["#application-form input", "#application-form textarea", "#application-form select"],
      containers: [".form-group", ".question", "fieldset"],
      labels: [".control-label", ".question-label", "legend"],
      options: [".select2-results__option", "[role='option']"]
    }),
    adapter("jobvite", "Jobvite", [/(^|\.)jobvite\.com$/i], {
      markers: [".jv-careersite", "[class*='jv-']"],
      controls: [".jv-careersite input", ".jv-careersite textarea", ".jv-careersite select", ".jv-careersite [role='combobox']"],
      containers: [".jv-form-field", ".form-group", "fieldset"],
      labels: [".jv-form-label", ".control-label", "legend"],
      options: [".jv-select-option", "[role='option']"]
    }),
    adapter("lever", "Lever", [/(^|\.)lever\.co$/i], {
      markers: [".application-form", ".application-question"],
      controls: [".application-form input", ".application-form textarea", ".application-form select", ".application-form [role='combobox']"],
      containers: [".application-question", ".application-field", ".form-group", "fieldset"],
      labels: [".application-label", ".application-question label", "legend"],
      options: [".application-dropdown [role='option']", "[class*='option']"]
    }),
    adapter("oracle", "Oracle Recruiting", [/(^|\.)oraclecloud\.com$/i], {
      markers: ["[data-oj-binding-provider]", "oj-input-text", "oj-select-single"],
      controls: ["oj-input-text input", "oj-text-area textarea", "oj-select-single input", "oj-combobox-one input", "[class*='oj-select'] [role='combobox']"],
      containers: ["oj-form-layout", ".oj-form-control", "[class*='field']", "fieldset"],
      labels: [".oj-label", "label.oj-label", "legend"],
      options: [".oj-listbox-result", "[role='option']"]
    }),
    adapter("successfactors", "SAP SuccessFactors", [/(^|\.)successfactors\.com$/i], {
      markers: ["[id*='careerSite']", "[class*='sapUiForm']", ".careerSite"],
      controls: [".careerSite input", ".careerSite textarea", ".careerSite select", "[class*='sapM'] [role='combobox']"],
      containers: [".sapUiFormElement", ".fieldWrapper", ".form-group", "fieldset"],
      labels: [".sapMLabel", ".sapUiFormLabel", "legend"],
      options: [".sapMSelectListItemBase", ".sapMComboBoxBaseItem", "[role='option']"]
    }),
    adapter("personio", "Personio", [/(^|\.)personio\.(de|com)$/i, /(^|\.)jobs\.personio\./i], {
      markers: ["[data-testid*='application-form']", "[class*='ApplicationForm']"],
      containers: ["[data-testid*='field']", "[class*='FormField']", "fieldset"],
      labels: ["[data-testid*='label']", "[class*='Label']", "legend"],
      options: ["[data-testid*='option']", "[role='option']"]
    }),
    adapter("rippling", "Rippling", [/(^|\.)rippling\.com$/i], {
      markers: ["[data-testid*='application']", "[class*='ApplicationForm']"],
      containers: ["[data-testid*='form-field']", "[class*='FormField']", "fieldset"],
      labels: ["[data-testid*='label']", "[class*='Label']", "legend"],
      options: ["[data-testid*='option']", "[role='option']"],
      uploads: ["input[data-testid*='resume']", "input[name*='resume' i]"]
    }),
    adapter("recruitee", "Recruitee", [/(^|\.)recruitee\.com$/i], {
      markers: ["[data-testid*='application-form']", ".application-form"],
      containers: ["[data-testid*='field']", ".form-field", ".question", "fieldset"],
      labels: ["[data-testid*='label']", ".form-label", "legend"],
      options: ["[data-testid*='option']", "[role='option']"]
    }),
    adapter("smartrecruiters", "SmartRecruiters", [/(^|\.)smartrecruiters\.com$/i], {
      markers: ["[data-test*='application']", "[class*='application-form']"],
      controls: ["[data-test*='application'] input", "[data-test*='application'] textarea", "[data-test*='application'] select", "[data-test*='application'] [role='combobox']"],
      containers: ["[data-test*='field']", "[class*='Field']", "fieldset"],
      labels: ["[data-test*='label']", "[class*='Label']", "legend"],
      options: ["[data-test*='option']", "[role='option']"]
    }),
    adapter("sourceflow", "Sourceflow", [/(^|\.)sourceflow\.co(\.uk)?$/i], {
      markers: ["[data-sourceflow]", "[class*='sourceflow']", "form[action*='application']"],
      containers: [".form-group", "[class*='form-field']", "fieldset"],
      labels: [".form-label", "[class*='field-label']", "legend"],
      options: ["[class*='select-option']", "[role='option']"]
    }),
    adapter("teamtailor", "Teamtailor", [/(^|\.)teamtailor\.com$/i], {
      markers: [
        "#job-application-form[data-controller~='careersite--form']",
        "[data-controller~='forms--inputs--upload']",
        "[data-controller*='jobs--application']",
        "[class*='application-form']"
      ],
      controls: ["[data-controller*='application'] input", "[data-controller*='application'] textarea", "[data-controller*='application'] select", "[data-controller*='application'] [role='combobox']"],
      containers: ["[data-question-id]", ".form-group", "[class*='field']", "fieldset"],
      labels: ["[data-question-label]", ".form-label", "legend"],
      options: ["[data-option-id]", "[role='option']"],
      uploads: [
        "[data-controller~='forms--inputs--upload'][id*='resume' i] input[type='file']",
        "[data-controller~='forms--inputs--upload'] input.dz-hidden-input",
        "input.dz-hidden-input"
      ]
    }),
    adapter("wellfound", "Wellfound", [/(^|\.)wellfound\.com$/i, /(^|\.)angel\.co$/i], {
      markers: ["[data-test*='application']", "[class*='Application']"],
      containers: ["[data-test*='field']", "[class*='FormField']", "fieldset"],
      labels: ["[data-test*='label']", "[class*='Label']", "legend"],
      options: ["[data-test*='option']", "[role='option']"]
    }),
    adapter("workable", "Workable", [/(^|\.)workable\.com$/i], {
      markers: ["[data-ui*='application']", "[data-testid*='application']", ".application-form"],
      controls: ["[data-ui*='application'] input", "[data-ui*='application'] textarea", "[data-ui*='application'] select", "[data-ui*='application'] [role='combobox']"],
      containers: ["[data-ui*='form-field']", "[data-testid*='field']", "[class*='fieldEntry']", "fieldset"],
      labels: ["[data-ui*='label']", "[data-testid*='label']", "legend"],
      options: ["[data-ui*='option']", "[data-testid*='option']", "[role='option']"]
    }),
    adapter("workday", "Workday", [/(^|\.)myworkdayjobs\.com$/i, /(^|\.)workday\.com$/i], {
      markers: ["[data-automation-id='jobApplicationPage']", "[data-automation-id='applicationForm']"],
      controls: ["[data-automation-id='jobApplicationPage'] input", "[data-automation-id='jobApplicationPage'] textarea", "[data-automation-id='jobApplicationPage'] select", "[data-automation-id='jobApplicationPage'] [role='combobox']", "[data-automation-id*='select']"],
      containers: ["[data-automation-id='formField']", "[data-automation-id*='question']", "[class*='css-'] fieldset", "fieldset"],
      labels: ["[data-automation-id='formLabel']", "[data-automation-id*='label']", "legend"],
      options: ["[data-automation-id='promptOption']", "[data-automation-id*='menuItem']", "[role='option']"],
      uploads: ["input[data-automation-id*='fileUpload']", "input[type='file']"]
    }),
    adapter("ycombinator", "Work at a Startup", [/(^|\.)workatastartup\.com$/i, /(^|\.)ycombinator\.com$/i], {
      markers: ["[data-testid*='application']", "form[action*='apply']"],
      containers: ["[data-testid*='field']", ".form-group", "fieldset"],
      labels: ["[data-testid*='label']", ".form-label", "legend"],
      options: ["[data-testid*='option']", "[role='option']"]
    })
  ];

  function adapter(id, name, hosts, configuration = {}) {
    return {
      id,
      name,
      hosts,
      markers: configuration.markers || [],
      controls: configuration.controls || [],
      containers: configuration.containers || [],
      labels: configuration.labels || [],
      options: configuration.options || [],
      uploads: configuration.uploads || []
    };
  }

  function create(helpers = {}) {
    const queryAll = helpers.queryAll || ((selector, root = document) => Array.from(root.querySelectorAll(selector)));
    const isVisible = helpers.isVisible || (() => true);
    const cleanText = helpers.cleanText || ((value) => String(value || "").replace(/\s+/g, " ").trim());
    let cachedKey = "";
    let cachedAdapter = null;

    function active() {
      const key = `${location.hostname}|${document.documentElement?.childElementCount || 0}`;
      if (cachedKey === key && cachedAdapter?.id !== "common") return cachedAdapter;
      cachedKey = key;
      cachedAdapter = detectAdapter(location.hostname, document, queryAll);
      return cachedAdapter;
    }

    function getControlSelectors() {
      return unique([...COMMON_CONTROL_SELECTORS, ...active().controls]);
    }

    function getOptionSelectors() {
      return unique([...COMMON_OPTION_SELECTORS, ...active().options]);
    }

    function getFieldContainer(control) {
      for (const selector of active().containers) {
        const container = safeClosest(control, selector);
        if (container) return container;
      }
      return null;
    }

    function getLabelCandidates(control) {
      const container = getFieldContainer(control) || control.parentElement;
      if (!container) return [];
      const candidates = [];
      for (const selector of active().labels) {
        for (const element of safeQueryAll(container, selector)) {
          if (element === control || element.contains?.(control)) continue;
          const text = cleanText(element.textContent || element.getAttribute?.("aria-label") || "");
          if (text) candidates.push(text);
        }
      }
      return unique(candidates);
    }

    function isRequired(control) {
      const container = getFieldContainer(control);
      const nestedControls = container
        ? safeQueryAll(container, "input, textarea, select, [role='checkbox'], [role='radio'], [role='combobox']").length
        : 0;
      const scopedContainer = nestedControls <= 1;
      const text = cleanText([
        control.getAttribute?.("aria-label"),
        control.getAttribute?.("data-automation-id"),
        ...getLabelCandidates(control),
        scopedContainer ? container?.textContent : ""
      ].filter(Boolean).join(" "));
      return Boolean(
        control.required ||
        control.getAttribute?.("aria-required") === "true" ||
        container?.matches?.("[aria-required='true'], [data-required='true']") ||
        scopedContainer && container?.querySelector?.("[aria-required='true'], [required], [class*='required' i]") ||
        /(^|\s)\*(\s|$)/.test(text) ||
        /\brequired\b/i.test(text)
      );
    }

    function getControlType(control) {
      if (control.matches?.("[contenteditable='true'][role='textbox']")) return "contenteditable";
      if (control.matches?.("[role='button'][aria-haspopup='listbox'], [role='button'][aria-haspopup='menu']")) return "combobox";
      if (control.matches?.(".fab-Select__control, [class*='Select__control'], .fab-SelectToggle, [data-fabric-component='SelectToggle'], .select2-selection, .select2-choice, .chosen-single")) return "combobox";
      return "";
    }

    function getVisibleOptions(control) {
      const options = [];
      for (const selector of getOptionSelectors()) {
        for (const option of queryAll(selector)) {
          if (option === control || option.contains?.(control) || !isVisible(option)) continue;
          const text = cleanText(option.textContent || option.getAttribute?.("aria-label") || option.getAttribute?.("data-value") || "");
          if (!text || text.length > 300) continue;
          options.push(option);
        }
      }
      return unique(options);
    }

    function getUploadSelectors() {
      return unique(active().uploads);
    }

    function describe() {
      const selected = active();
      return { id: selected.id, name: selected.name };
    }

    return {
      describe,
      getControlSelectors,
      getControlType,
      getFieldContainer,
      getLabelCandidates,
      getOptionSelectors,
      getUploadSelectors,
      getVisibleOptions,
      isRequired
    };
  }

  function detectAdapter(hostname, root, queryAll) {
    const host = String(hostname || "").toLowerCase();
    const hostMatch = DEFINITIONS.find((definition) => definition.hosts.some((pattern) => pattern.test(host)));
    if (hostMatch) return hostMatch;

    return DEFINITIONS.find((definition) => definition.markers.some((selector) => {
      try {
        return queryAll(selector, root).length > 0;
      } catch (_error) {
        return false;
      }
    })) || adapter("common", "Common form", [], {});
  }

  function safeClosest(element, selector) {
    try {
      return element?.closest?.(selector) || null;
    } catch (_error) {
      return null;
    }
  }

  function safeQueryAll(root, selector) {
    try {
      return Array.from(root?.querySelectorAll?.(selector) || []);
    } catch (_error) {
      return [];
    }
  }

  function unique(values) {
    return Array.from(new Set(values.filter(Boolean)));
  }

  window.AutoBidAtsAdapters = {
    create,
    supported: DEFINITIONS.map(({ id, name }) => ({ id, name }))
  };
})();
