import 'bluebird-global'
import botbuilder, { ActivityTypes, BotFrameworkAdapter, TurnContext } from 'botbuilder'
import * as sdk from 'botpress/sdk'

const channelMicrosoft = 'channel-microsoft'

const onServerStarted = async (bp: typeof sdk) => {}

const onServerReady = async (bp: typeof sdk) => {
  const channelRouter = bp.http.createRouterForBot(channelMicrosoft, { checkAuthentication: false })

  channelRouter.post('/api/messages', async (req, res) => {
    const botId = req.params.botId
    const config = await bp.config.getModuleConfigForBot(channelMicrosoft, botId)
    console.log('config', config)

    const adapter = new BotFrameworkAdapter({
      appId: config.microsoftAppId,
      appPassword: config.microsoftAppPassword
    })

    adapter.onTurnError = async (context, error) => {
      bp.logger
        .forBot(botId)
        .attachError(error)
        .error(error.message)
      await context.sendActivity(`Oops. Something went wrong!`)
    }

    await adapter.processActivity(req, res, async (context: TurnContext) => {
      if (context.activity.type === ActivityTypes.Message) {
        const message = context.activity.text
        const accountId = context.activity.from.id
        const content = await bp.converse.sendMessage(botId, accountId, { text: message }, 'microsoft')

        for (let i = 0; i < content.responses.length; i += 2) {
          const isTyping = content.responses[i].value
          const message = content.responses[i + 1]

          // Format of message with typing indicator
          // [
          //   { type: 'typing' },
          //   { type: 'delay', value: 2000 },
          //   { type: 'message', text: 'Hello... How are you?' }
          // ]

          let activities = [message]
          if (isTyping) {
            activities = [{ type: 'typing' }, { type: 'delay', value: 250 }, message]
          }

          await context.sendActivities(activities)
        }
      }
    })
  })

  channelRouter.get('/', (req, res) => {
    res.sendStatus(200)
  })
}

const entryPoint: sdk.ModuleEntryPoint = {
  onServerStarted,
  onServerReady,
  definition: {
    name: channelMicrosoft,
    fullName: 'Microsoft Channel',
    homepage: 'https://botpress.io',
    noInterface: true
  }
}

export default entryPoint
