/** Base class for every error `@uekichinos/quire` throws — catch this to catch them all. */
export class QuireError extends Error {
  constructor(message: string) {
    super(message.startsWith('@uekichinos/quire') ? message : `@uekichinos/quire: ${message}`)
    this.name = 'QuireError'
  }
}
