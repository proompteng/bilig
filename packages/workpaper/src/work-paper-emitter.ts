import type { WorkPaperDetailedEventMap, WorkPaperDetailedListener, WorkPaperEventName, WorkPaperListener } from './work-paper-types.js'

type ListenerMap = {
  [EventName in WorkPaperEventName]: Set<WorkPaperListener<EventName>>
}

type DetailedListenerMap = {
  [EventName in WorkPaperEventName]: Set<WorkPaperDetailedListener<EventName>>
}

export type WorkPaperDetailedEvent = {
  [EventName in WorkPaperEventName]: {
    eventName: EventName
    payload: WorkPaperDetailedEventMap[EventName]
  }
}[WorkPaperEventName]

export class WorkPaperEmitter {
  private listenerCount = 0
  private detailedListenerCount = 0

  private readonly listeners: ListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  }

  private readonly detailedListeners: DetailedListenerMap = {
    sheetAdded: new Set(),
    sheetRemoved: new Set(),
    sheetRenamed: new Set(),
    namedExpressionAdded: new Set(),
    namedExpressionRemoved: new Set(),
    valuesUpdated: new Set(),
    evaluationSuspended: new Set(),
    evaluationResumed: new Set(),
  }

  on<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    const listeners = this.listeners[eventName]
    if (!listeners.has(listener)) {
      listeners.add(listener)
      this.listenerCount += 1
    }
  }

  off<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    if (this.listeners[eventName].delete(listener)) {
      this.listenerCount -= 1
    }
  }

  once<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperListener<EventName>): void {
    const wrapper: WorkPaperListener<EventName> = (...args) => {
      this.off(eventName, wrapper)
      listener(...args)
    }
    this.on(eventName, wrapper)
  }

  onDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    const listeners = this.detailedListeners[eventName]
    if (!listeners.has(listener)) {
      listeners.add(listener)
      this.detailedListenerCount += 1
    }
  }

  offDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    if (this.detailedListeners[eventName].delete(listener)) {
      this.detailedListenerCount -= 1
    }
  }

  onceDetailed<EventName extends WorkPaperEventName>(eventName: EventName, listener: WorkPaperDetailedListener<EventName>): void {
    const wrapper: WorkPaperDetailedListener<EventName> = (payload) => {
      this.offDetailed(eventName, wrapper)
      listener(payload)
    }
    this.onDetailed(eventName, wrapper)
  }

  hasListeners(eventName: WorkPaperEventName): boolean {
    return this.listeners[eventName].size > 0 || this.detailedListeners[eventName].size > 0
  }

  get anyListenerCount(): number {
    return this.listenerCount + this.detailedListenerCount
  }

  hasAnyListeners(): boolean {
    return this.listenerCount > 0 || this.detailedListenerCount > 0
  }

  emitDetailed(event: WorkPaperDetailedEvent): void {
    this.dispatchDetailed(event)
  }

  private dispatchDetailed(event: WorkPaperDetailedEvent): void {
    switch (event.eventName) {
      case 'sheetAdded':
        this.listeners.sheetAdded.forEach((listener) => {
          listener(event.payload.sheetName)
        })
        this.detailedListeners.sheetAdded.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'sheetRemoved':
        this.listeners.sheetRemoved.forEach((listener) => {
          listener(event.payload.sheetName, event.payload.changes)
        })
        this.detailedListeners.sheetRemoved.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'sheetRenamed':
        this.listeners.sheetRenamed.forEach((listener) => {
          listener(event.payload.oldName, event.payload.newName)
        })
        this.detailedListeners.sheetRenamed.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'namedExpressionAdded':
        this.listeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload.name, event.payload.changes)
        })
        this.detailedListeners.namedExpressionAdded.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'namedExpressionRemoved':
        this.listeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload.name, event.payload.changes)
        })
        this.detailedListeners.namedExpressionRemoved.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'valuesUpdated':
        this.listeners.valuesUpdated.forEach((listener) => {
          listener(event.payload.changes)
        })
        this.detailedListeners.valuesUpdated.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'evaluationSuspended':
        this.listeners.evaluationSuspended.forEach((listener) => {
          listener()
        })
        this.detailedListeners.evaluationSuspended.forEach((listener) => {
          listener(event.payload)
        })
        return
      case 'evaluationResumed':
        this.listeners.evaluationResumed.forEach((listener) => {
          listener(event.payload.changes)
        })
        this.detailedListeners.evaluationResumed.forEach((listener) => {
          listener(event.payload)
        })
    }
  }

  clear(): void {
    Object.values(this.listeners).forEach((listeners) => listeners.clear())
    Object.values(this.detailedListeners).forEach((listeners) => listeners.clear())
    this.listenerCount = 0
    this.detailedListenerCount = 0
  }
}
