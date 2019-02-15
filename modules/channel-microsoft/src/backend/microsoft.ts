import { Activity, ActivityTypes, BotFrameworkAdapter, CardFactory, MessageFactory, TurnContext } from 'botbuilder'
import * as sdk from 'botpress/sdk'
import mime from 'mime/lite'

import { Config } from '../config'

export const channelMicrosoft = 'channel-microsoft'

export class MicrosoftService {
  private microsoftBots: { [key: string]: ScopedMicrosoftService } = {}

  constructor(private bp: typeof sdk) {}

  async init() {
    const channelRouter = this.bp.http.createRouterForBot(channelMicrosoft, { checkAuthentication: false })

    channelRouter.post('/api/messages', async (req, res) => {
      const botId = req.params.botId
      const microsoftService = await this.forBot(botId)
      await microsoftService.processRequest(req, res)
    })

    channelRouter.get('/', (req, res) => {
      res.sendStatus(200)
    })
  }

  async forBot(botId: string): Promise<ScopedMicrosoftService> {
    if (this.microsoftBots[botId]) {
      return this.microsoftBots[botId]
    }

    const config = await this.bp.config.getModuleConfigForBot(channelMicrosoft, botId)
    this.microsoftBots[botId] = new ScopedMicrosoftService(this.bp, botId, config)

    return this.microsoftBots[botId]
  }
}

export class ScopedMicrosoftService {
  private adapter: BotFrameworkAdapter

  constructor(private bp: typeof sdk, private botId: string, private config: Config) {
    this.adapter = new BotFrameworkAdapter({
      appId: this.config.microsoftAppId,
      appPassword: this.config.microsoftAppPassword
    })

    this.adapter.onTurnError = async (context, error) => {
      this.bp.logger
        .forBot(botId)
        .attachError(error)
        .error(error.message)
      await context.sendActivity(`Oops. Something went wrong! Please try something else.`)
    }
  }

  async processRequest(req, res): Promise<void> {
    await this.adapter.processActivity(req, res, async (context: TurnContext) => {
      if (context.activity.type === ActivityTypes.Message) {
        const message = context.activity.text
        const accountId = context.activity.from.id
        const content = await this.bp.converse.sendMessage(this.botId, accountId, { text: message }, 'microsoft')

        for (let i = 0; i < content.responses.length; i += 2) {
          const isTyping = content.responses[i].value
          const message = content.responses[i + 1]
          const activities = []

          if (isTyping) {
            activities.push({ type: 'typing' }, { type: 'delay', value: 250 })
          }

          if (message.actions) {
            activities.push(this._handleActions(message))
          } else if (message.attachments) {
            activities.push(this._handleAttachments(message))
          } else if (message.image) {
            activities.push(this._handleImage(message))
          } else if (message.text) {
            activities.push(message)
          } else {
            throw new Error('channel-microsoft does not recognize any content type')
          }

          await context.sendActivities(activities)
        }
      }
    })
  }

  private _handleActions(message): Partial<Activity> {
    return MessageFactory.suggestedActions(message.actions, message.text)
  }

  private _handleAttachments(message): Partial<Activity> {
    // See hero card data reference:
    // https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/cards/cards-reference#hero-card
    const attachments = message.attachments.map(a => {
      return CardFactory.heroCard(a.title, a.subtitle, a.images, a.buttons)
    })
    return MessageFactory.carousel(attachments)
  }

  private _handleImage(message): Partial<Activity> {
    return MessageFactory.contentUrl(message.image.url, mime.getType(message.image.url), message.image.title)
  }
}
