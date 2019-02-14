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

        console.log('message', message)
        console.log('accountId', accountId)

        const content = await bp.converse.sendMessage(botId, accountId, { text: message }, 'microsoft')
        for (let i = 0; i < content.responses.length; i += 2) {
          const typing = content.responses[i]
          const message = content.responses[i + 1]
          await context.sendActivity(message.text)
        }
      }
    })
  })

  channelRouter.get('/', (req, res) => {
    res.sendStatus(200)
  })
}

const onBotMount = async (bp: typeof sdk, botId: string) => {}

const entryPoint: sdk.ModuleEntryPoint = {
  onServerStarted,
  onServerReady,
  onBotMount,
  definition: {
    name: channelMicrosoft,
    fullName: 'Microsoft Channel',
    homepage: 'https://botpress.io',
    noInterface: true
  }
}

export default entryPoint
