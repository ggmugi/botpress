import { ActivityTypes, BotFrameworkAdapter, MessageFactory, TurnContext } from 'botbuilder'
import * as sdk from 'botpress/sdk'

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
      await context.sendActivity(`Oops. Something went wrong!`)
    }
  }

  async processRequest(req, res): Promise<void> {
    await this.adapter.processActivity(req, res, async (context: TurnContext) => {
      if (context.activity.type === ActivityTypes.Message) {
        const message = context.activity.text
        const accountId = context.activity.from.id
        const content = await this.bp.converse.sendMessage(this.botId, accountId, { text: message }, 'microsoft')

        for (let i = 0; i < content.responses.length; i += 2) {
          // NOTE: Microsoft's message format w/ typing indicator
          // [
          //   { type: 'typing' },
          //   { type: 'delay', value: 2000 },
          //   { type: 'message', text: 'Hello... How are you?' }
          // ]
          const isTyping = content.responses[i].value
          const message = content.responses[i + 1]
          let activities = [message]

          if (isTyping) {
            activities = [{ type: 'typing' }, { type: 'delay', value: 250 }]
          }

          if (message.actions) {
            const activity = MessageFactory.suggestedActions(message.actions, message.text)
            activities = [...activities, activity]
          } else if (message.text) {
            activities = [...activities, message]
          }

          await context.sendActivities(activities)
        }
      }
    })
  }
}
