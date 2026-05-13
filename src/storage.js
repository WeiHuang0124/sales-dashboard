// Drop-in replacement for window.storage using localStorage
// Same async API so App.jsx works without changes

const storage = {
  async get(key) {
    try {
      const value = localStorage.getItem(key)
      return value !== null ? { key, value } : null
    } catch { return null }
  },
  async set(key, value) {
    try {
      localStorage.setItem(key, value)
      return { key, value }
    } catch { return null }
  },
  async delete(key) {
    try {
      localStorage.removeItem(key)
      return { key, deleted: true }
    } catch { return null }
  },
  async list(prefix = '') {
    try {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(prefix))
      return { keys, prefix }
    } catch { return { keys: [] } }
  },
}

export default storage
