(function registerDevTools(globalScope) {
  if (globalScope.pbDevPro) return;

  globalScope.pbDevPro = {
    async on() {
      await chrome.storage.local.set({
        pbDeveloperProOverride: true,
        pbDeveloperProOverrideAt: Date.now()
      });
      console.log('PlugBits Developer Pro ENABLED');
    },

    async off() {
      await chrome.storage.local.remove([
        'pbDeveloperProOverride',
        'pbDeveloperProOverrideAt'
      ]);
      console.log('PlugBits Developer Pro DISABLED');
    },

    async status() {
      const raw = await chrome.storage.local.get([
        'pbDeveloperProOverride',
        'pbDeveloperProOverrideAt'
      ]);
      const result = {
        enabled: Boolean(raw.pbDeveloperProOverride),
        pbDeveloperProOverride: raw.pbDeveloperProOverride ?? false,
        pbDeveloperProOverrideAt: raw.pbDeveloperProOverrideAt ?? 0
      };
      console.log(result);
      return result;
    }
  };
})(globalThis);