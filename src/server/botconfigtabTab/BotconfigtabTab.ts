import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/botconfigtabTab/index.html")
@PreventIframe("/botconfigtabTab/config.html")
@PreventIframe("/botconfigtabTab/remove.html")
export class BotconfigtabTab {
}
