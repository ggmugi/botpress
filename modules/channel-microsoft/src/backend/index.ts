import 'bluebird-global'
import * as sdk from 'botpress/sdk'

import { channelMicrosoft, MicrosoftService } from './microsoft'

const onServerStarted = async (bp: typeof sdk) => {}

const onServerReady = async (bp: typeof sdk) => {
  const microsoft = new MicrosoftService(bp)
  await microsoft.init()
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
