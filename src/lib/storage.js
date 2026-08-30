// Minimal localStorage-backed shim for the window.storage.get/set API the
// app was originally written against (that API is only available inside
// Claude Artifacts). Same async shape, backed by the browser instead.
function installLocalStorageShim() {
  if (typeof window === "undefined" || window.storage) return;

  window.storage = {
    async get(key) {
      try {
        const value = window.localStorage.getItem(key);
        return value === null ? null : { value };
      } catch (e) {
        throw new Error(e && e.message ? e.message : "No se pudo leer de localStorage.");
      }
    },
    async set(key, value) {
      try {
        window.localStorage.setItem(key, value);
        return true;
      } catch (e) {
        throw new Error(e && e.message ? e.message : "No se pudo escribir en localStorage.");
      }
    },
  };
}

export default installLocalStorageShim;
