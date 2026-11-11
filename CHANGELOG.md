# Changelog

## [0.12.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.11.0...v0.12.0) (2026-11-11)


### Features

* **mail:** write and read Markdown bodies ([#130](https://github.com/chrischall/office-outlook-mcp/issues/130)) ([5161906](https://github.com/chrischall/office-outlook-mcp/commit/5161906e5ba287e729a664b19ded5a86b8eb5cd6))
* **reply:** add cc and bcc to replies, and addCc/removeCc to draft edits ([#128](https://github.com/chrischall/office-outlook-mcp/issues/128)) ([08b692b](https://github.com/chrischall/office-outlook-mcp/commit/08b692bde6080da5e002e61ba455b1c2cb31d180))


### Bug Fixes

* **mail:** keep html: true returning the raw HTML body ([#132](https://github.com/chrischall/office-outlook-mcp/issues/132)) ([9586c8d](https://github.com/chrischall/office-outlook-mcp/commit/9586c8da162e25c5525078b33ab6b03f9ef2d002))

## [0.11.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.10.1...v0.11.0) (2026-11-10)


### Features

* **drafts:** edit unsent drafts with outlook_update_draft ([#126](https://github.com/chrischall/office-outlook-mcp/issues/126)) ([58012a3](https://github.com/chrischall/office-outlook-mcp/commit/58012a31a3b14caf43f692794946ceda65bb1018))
* **drafts:** send an existing draft with outlook_send_draft ([#127](https://github.com/chrischall/office-outlook-mcp/issues/127)) ([8ec9d82](https://github.com/chrischall/office-outlook-mcp/commit/8ec9d8209d21765aca8870b8b9e8447d7700c62e))
* **folders:** create, rename, move and delete mail folders ([#121](https://github.com/chrischall/office-outlook-mcp/issues/121)) ([b4711d1](https://github.com/chrischall/office-outlook-mcp/commit/b4711d1bbee19b64117d5695be586ef91cb5544b))


### Bug Fixes

* **folders:** accurate errors and previews, and safe large-folder deletes ([#124](https://github.com/chrischall/office-outlook-mcp/issues/124)) ([a17a232](https://github.com/chrischall/office-outlook-mcp/commit/a17a232b81ba468c97d9c358c1d968c93bdf7266))
* **folders:** count subfolder mail in delete confirmations ([#125](https://github.com/chrischall/office-outlook-mcp/issues/125)) ([963a0a3](https://github.com/chrischall/office-outlook-mcp/commit/963a0a32aa39b14d4a16d4d1022401ef719de9d4))

## [0.10.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.10.0...v0.10.1) (2026-11-09)


### Bug Fixes

* **deps:** bump @fetchproxy/server to 3.8 and MCP SDK to 2.4 ([#118](https://github.com/chrischall/office-outlook-mcp/issues/118)) ([75f678e](https://github.com/chrischall/office-outlook-mcp/commit/75f678e568fec465b6c9419386258964985161f0))

## [0.10.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.9.0...v0.10.0) (2026-11-09)


### Features

* **copilot:** label times in the mailbox zone, add verbatim mode, protect code indentation ([#115](https://github.com/chrischall/office-outlook-mcp/issues/115)) ([d883501](https://github.com/chrischall/office-outlook-mcp/commit/d8835019ca389d0761f496dfb5aa6d221c2dde3a))


### Bug Fixes

* **calendar:** return attendees from outlook_get_event in every view ([#114](https://github.com/chrischall/office-outlook-mcp/issues/114)) ([5d7137d](https://github.com/chrischall/office-outlook-mcp/commit/5d7137d7606affb8f0e8abec0a73b4d2afdb03c3))

## [0.9.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.8.4...v0.9.0) (2026-11-09)


### Features

* **copilot:** ask Microsoft 365 Copilot with your mailbox as context ([#113](https://github.com/chrischall/office-outlook-mcp/issues/113)) ([604d266](https://github.com/chrischall/office-outlook-mcp/commit/604d2663a29ebcdd89150c63a548018d9268808c))
* **rules:** manage inbox rules and automatic replies ([#111](https://github.com/chrischall/office-outlook-mcp/issues/111)) ([11f9fd9](https://github.com/chrischall/office-outlook-mcp/commit/11f9fd93f72d20da5ab45bbffbc94ef06aae2858))

## [0.8.4](https://github.com/chrischall/office-outlook-mcp/compare/v0.8.3...v0.8.4) (2026-11-07)


### Bug Fixes

* **attachments:** hand over PDFs, Office files and large attachments instead of metadata only ([#107](https://github.com/chrischall/office-outlook-mcp/issues/107)) ([5b0ed71](https://github.com/chrischall/office-outlook-mcp/commit/5b0ed71d19ba5534127e7afdb542e692885acf2c))
* **deps:** upgrade @chrischall/mcp-utils to 3.7.2 so uploads carry only the file's bytes ([#109](https://github.com/chrischall/office-outlook-mcp/issues/109)) ([a2c0c85](https://github.com/chrischall/office-outlook-mcp/commit/a2c0c85aba64057d1eae69a4a8f97b9a914da539))

## [0.8.3](https://github.com/chrischall/office-outlook-mcp/compare/v0.8.2...v0.8.3) (2026-11-06)


### Bug Fixes

* **paging:** keep next_from and export cursors short when many messages share a second ([#105](https://github.com/chrischall/office-outlook-mcp/issues/105)) ([6834acd](https://github.com/chrischall/office-outlook-mcp/commit/6834acdef8cfd2f358b71fbd1becd6106d53ba79))

## [0.8.2](https://github.com/chrischall/office-outlook-mcp/compare/v0.8.1...v0.8.2) (2026-11-06)


### Bug Fixes

* **auth:** stop opening a new Outlook tab on every capture ([#103](https://github.com/chrischall/office-outlook-mcp/issues/103)) ([5c9abb4](https://github.com/chrischall/office-outlook-mcp/commit/5c9abb49b8f4e0aeaea798f040d93e5df21d495e))

## [0.8.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.8.0...v0.8.1) (2026-11-04)


### Bug Fixes

* **brief:** align the brief view with the fleet contract ([#101](https://github.com/chrischall/office-outlook-mcp/issues/101)) ([4a5723a](https://github.com/chrischall/office-outlook-mcp/commit/4a5723a7c03d54b3add4cfd6b3d60e4c2e8ec7b5))

## [0.8.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.7.1...v0.8.0) (2026-11-04)


### Features

* context-group hints, a brief view on unread mail, the calendar and tasks ([#98](https://github.com/chrischall/office-outlook-mcp/issues/98)) ([d0c606b](https://github.com/chrischall/office-outlook-mcp/commit/d0c606be26fe5e9483df12d2df5dd97bd6b4b2ea))


### Bug Fixes

* **deps:** bump @fetchproxy/server from 3.6.2 to 3.6.3 in the production-dependencies group ([#94](https://github.com/chrischall/office-outlook-mcp/issues/94)) ([cafc239](https://github.com/chrischall/office-outlook-mcp/commit/cafc2395231bc8c016460b324b9f09479973647f))
* **deps:** bump the production-dependencies group across 1 directory with 3 updates ([#97](https://github.com/chrischall/office-outlook-mcp/issues/97)) ([29d57cd](https://github.com/chrischall/office-outlook-mcp/commit/29d57cd16bf2e6e544ac6876bb5464ce4f020978))
* fence brief mail, calendar and task lines and note a further page ([#100](https://github.com/chrischall/office-outlook-mcp/issues/100)) ([7377f88](https://github.com/chrischall/office-outlook-mcp/commit/7377f88e15d290d42a2e38a1e922f0a9ab385ecf))

## [0.7.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.7.0...v0.7.1) (2026-10-21)


### Bug Fixes

* **deps:** bump @chrischall/mcp-utils to 3.3.1 so a server error after sendmail is outcome-unknown ([#91](https://github.com/chrischall/office-outlook-mcp/issues/91)) ([e31e1df](https://github.com/chrischall/office-outlook-mcp/commit/e31e1df0a8d287bfdfb7244c24fe3f44468a0736))

## [0.7.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.6.2...v0.7.0) (2026-10-20)


### ⚠ BREAKING CHANGES

* **tools:** rename tools to the list/get naming rule ([#89](https://github.com/chrischall/office-outlook-mcp/issues/89))

### Features

* **export:** bulk-export mail and calendar to local files or a download link ([#90](https://github.com/chrischall/office-outlook-mcp/issues/90)) ([0bf92a8](https://github.com/chrischall/office-outlook-mcp/commit/0bf92a834fb3dd8e1a7ed84c1de595c5749f40dd))
* **tools:** rename tools to the list/get naming rule ([#89](https://github.com/chrischall/office-outlook-mcp/issues/89)) ([b4d017b](https://github.com/chrischall/office-outlook-mcp/commit/b4d017b24efe629eb74e309fc6a243ade2f76024))


### Bug Fixes

* **deps:** bump dotenv from 18.0.6 to 18.0.8 in the production-dependencies group ([#83](https://github.com/chrischall/office-outlook-mcp/issues/83)) ([26a1c4f](https://github.com/chrischall/office-outlook-mcp/commit/26a1c4fe3c6451d68d928f1174a8c1b17da32503))


### Documentation

* **mint:** add a note to the tasks sample prompt ([#88](https://github.com/chrischall/office-outlook-mcp/issues/88)) ([278ffb6](https://github.com/chrischall/office-outlook-mcp/commit/278ffb6561e0bc9b75d2ceca22ee56c7f2e239cb))
* **mint:** add sample prompts for the public catalog ([#86](https://github.com/chrischall/office-outlook-mcp/issues/86)) ([adfd112](https://github.com/chrischall/office-outlook-mcp/commit/adfd1124eea28ca60ae84167f870d372cd7cb6f1))

## [0.6.2](https://github.com/chrischall/office-outlook-mcp/compare/v0.6.1...v0.6.2) (2026-10-17)


### Bug Fixes

* **client:** name the host and cause when Outlook cannot be reached ([#79](https://github.com/chrischall/office-outlook-mcp/issues/79)) ([2c58250](https://github.com/chrischall/office-outlook-mcp/commit/2c582505879f31cc479b16845e2196006e17c0f1))

## [0.6.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.6.0...v0.6.1) (2026-10-16)


### Bug Fixes

* **calendar:** stop outlook_update_event turning meetings into Teams meetings by default ([#74](https://github.com/chrischall/office-outlook-mcp/issues/74)) ([00229e7](https://github.com/chrischall/office-outlook-mcp/commit/00229e71c0d682de9f5796e3e089f02bc7278f50))
* **plugin:** launch the published package so marketplace installs start ([#77](https://github.com/chrischall/office-outlook-mcp/issues/77)) ([9d86d8c](https://github.com/chrischall/office-outlook-mcp/commit/9d86d8cb2240d8fc7942ac3abe364d00e05f270c))

## [0.6.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.5.2...v0.6.0) (2026-10-14)


### Features

* **mail:** page mail oldest first from a UTC time with a resumable next_from ([#72](https://github.com/chrischall/office-outlook-mcp/issues/72)) ([d0d63da](https://github.com/chrischall/office-outlook-mcp/commit/d0d63daca949c973c976e6a029e979e9de82a4a6))

## [0.5.2](https://github.com/chrischall/office-outlook-mcp/compare/v0.5.1...v0.5.2) (2026-10-11)


### Bug Fixes

* **deps:** bump the production-dependencies group with 2 updates ([#68](https://github.com/chrischall/office-outlook-mcp/issues/68)) ([536115b](https://github.com/chrischall/office-outlook-mcp/commit/536115b6e5741a606aee7a2d4d7f6324cd60046a))

## [0.5.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.5.0...v0.5.1) (2026-10-10)


### Bug Fixes

* **mail:** tell agents an answered invite is already filed and tidy triage edge cases ([#65](https://github.com/chrischall/office-outlook-mcp/issues/65)) ([55e5131](https://github.com/chrischall/office-outlook-mcp/commit/55e5131aabc6a37871d3ae234d991107628fbe2e))

## [0.5.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.4.1...v0.5.0) (2026-10-10)


### Features

* triage unread mail, answer invites, reply, and manage calendar events ([#61](https://github.com/chrischall/office-outlook-mcp/issues/61)) ([09e9a61](https://github.com/chrischall/office-outlook-mcp/commit/09e9a61524ec93b83842895a14d190e5d4454e68))


### Documentation

* **mail:** say where the 20-address recipient cap applies and how to get every address ([#64](https://github.com/chrischall/office-outlook-mcp/issues/64)) ([24726b2](https://github.com/chrischall/office-outlook-mcp/commit/24726b22e96945551e61134c41f98071671f6b1e))

## [0.4.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.4.0...v0.4.1) (2026-10-09)


### Bug Fixes

* annotate tools truthfully and sync manifests with the served tools ([#58](https://github.com/chrischall/office-outlook-mcp/issues/58)) ([3753d9f](https://github.com/chrischall/office-outlook-mcp/commit/3753d9f336d230cfd2d45cdc582cba89334632d3))
* build Outlook query strings with mcp-utils' buildQueryString ([#56](https://github.com/chrischall/office-outlook-mcp/issues/56)) ([f1904c6](https://github.com/chrischall/office-outlook-mcp/commit/f1904c68e8cff040da11176c05b53dc6c1781061))
* **deps:** update @chrischall/mcp-utils to 3.0.0 ([#57](https://github.com/chrischall/office-outlook-mcp/issues/57)) ([c86e8c2](https://github.com/chrischall/office-outlook-mcp/commit/c86e8c296ea1419220fe497187862d763d2883c3))
* drop the esbuild defines that only silenced a lint false positive ([#59](https://github.com/chrischall/office-outlook-mcp/issues/59)) ([73360d5](https://github.com/chrischall/office-outlook-mcp/commit/73360d5f36916c5a71bb892c8641c28114e3739d))
* point the plugin at its MCP config with the mcpServers key Claude Code reads ([#60](https://github.com/chrischall/office-outlook-mcp/issues/60)) ([b00cc67](https://github.com/chrischall/office-outlook-mcp/commit/b00cc672f2f3756a81878569c3f66fffac2a49b5))
* resolve low-severity audit findings ([#54](https://github.com/chrischall/office-outlook-mcp/issues/54)) ([b0d4e86](https://github.com/chrischall/office-outlook-mcp/commit/b0d4e861d1f2aef519b16dca051bd96cc3be95df))

## [0.4.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.5...v0.4.0) (2026-10-08)


### Features

* **calendar:** find meeting times, and create and update Teams meetings ([#50](https://github.com/chrischall/office-outlook-mcp/issues/50)) ([6a233ee](https://github.com/chrischall/office-outlook-mcp/commit/6a233eef9cb0c97a4ea0603c3d7f41c1842defe6))


### Bug Fixes

* **calendar:** verify offset end times correctly when updating a meeting ([#53](https://github.com/chrischall/office-outlook-mcp/issues/53)) ([988d24a](https://github.com/chrischall/office-outlook-mcp/commit/988d24afed1fdf7979f5eb389861b0dfecf934e6))

## [0.3.5](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.4...v0.3.5) (2026-10-07)


### Bug Fixes

* **deps:** bump mcp-utils to 2.15.0 and fetchproxy to 3.6.0 for elicitation opt-out and relay frame fixes ([#46](https://github.com/chrischall/office-outlook-mcp/issues/46)) ([54a350e](https://github.com/chrischall/office-outlook-mcp/commit/54a350e3f317791b5e3c98e6b837cd9a63a52c57))
* **deps:** bump source-map-js from 1.2.1 to 1.2.2 in the security group across 1 directory ([#48](https://github.com/chrischall/office-outlook-mcp/issues/48)) ([dd5fe7a](https://github.com/chrischall/office-outlook-mcp/commit/dd5fe7acaf00ff93d7074544af63c0cf5c4f8709))


### Documentation

* document MCP_CONFIRM_ELICITATION ([#49](https://github.com/chrischall/office-outlook-mcp/issues/49)) ([58f9e9a](https://github.com/chrischall/office-outlook-mcp/commit/58f9e9a6458de31768b4fcbab61fffcfdc7503d5))

## [0.3.4](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.3...v0.3.4) (2026-10-05)


### Bug Fixes

* **deps:** bump the production-dependencies group with 2 updates ([#43](https://github.com/chrischall/office-outlook-mcp/issues/43)) ([a964fae](https://github.com/chrischall/office-outlook-mcp/commit/a964fae7b441f0231bd04008bcb389d1c6b6b5b6))
* **deps:** require @chrischall/mcp-utils 2.14.0 and MCP SDK 2.3.0 ([#45](https://github.com/chrischall/office-outlook-mcp/issues/45)) ([0b0d89c](https://github.com/chrischall/office-outlook-mcp/commit/0b0d89c4ed51c4587e44ed58e47d7c1e75bf0c03))

## [0.3.3](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.2...v0.3.3) (2026-10-03)


### Bug Fixes

* **deps:** adopt @chrischall/mcp-utils 2.12.0 untrusted framing and confirmWrite ([#39](https://github.com/chrischall/office-outlook-mcp/issues/39)) ([2573c5b](https://github.com/chrischall/office-outlook-mcp/commit/2573c5ba33ad4920e5417a21015385c990e618ae))
* **deps:** bump @chrischall/mcp-utils to 2.13.0 ([#40](https://github.com/chrischall/office-outlook-mcp/issues/40)) ([3bc7c4f](https://github.com/chrischall/office-outlook-mcp/commit/3bc7c4fb3cbe4cdef414a662b58df1959ff15c1a))
* keep credentials and report edge_blocked on CDN/WAF blocks (mcp-utils 2.10.0) ([#37](https://github.com/chrischall/office-outlook-mcp/issues/37)) ([6b15420](https://github.com/chrischall/office-outlook-mcp/commit/6b1542035baf6f88cb51cc460b582709cab83a97))
* keep write approvals valid across a hosted restart (mcp-utils 2.11.0) ([#38](https://github.com/chrischall/office-outlook-mcp/issues/38)) ([4165d51](https://github.com/chrischall/office-outlook-mcp/commit/4165d511feb0c2ac3e1db28427a07e09f34cd2f3))
* report CDN/WAF blocks as edge_blocked, not a rejected credential (mcp-utils 2.9.0) ([#34](https://github.com/chrischall/office-outlook-mcp/issues/34)) ([7464170](https://github.com/chrischall/office-outlook-mcp/commit/7464170580c8a326cd298b1d2566126ab1a8d5a1))

## [0.3.2](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.1...v0.3.2) (2026-09-27)


### Bug Fixes

* **deps:** bump the production-dependencies group with 3 updates ([#28](https://github.com/chrischall/office-outlook-mcp/issues/28)) ([dc2bc1a](https://github.com/chrischall/office-outlook-mcp/commit/dc2bc1a5c002da383df5c91916700bea063a3fa7))
* **deps:** move to [@fetchproxy](https://github.com/fetchproxy) 3.4 for ContextMint Bridge errors, capability subsets and managed pins ([#30](https://github.com/chrischall/office-outlook-mcp/issues/30)) ([aad986e](https://github.com/chrischall/office-outlook-mcp/commit/aad986e956681e4bddb3bf5d473506965bb8b052))
* **deps:** move to @chrischall/mcp-utils 2.8 and [@fetchproxy](https://github.com/fetchproxy) 3.4.1 for clearer browser-bridge errors ([#33](https://github.com/chrischall/office-outlook-mcp/issues/33)) ([89a7513](https://github.com/chrischall/office-outlook-mcp/commit/89a751344ff8120ad8e2d9206c2d41caca0dbb36))


### Documentation

* say extension and fpx match by protocol number, and where ContextMint Bridge comes from ([#32](https://github.com/chrischall/office-outlook-mcp/issues/32)) ([cef6ffc](https://github.com/chrischall/office-outlook-mcp/commit/cef6ffc42108a6334c2dd74943cb82753636cef3))

## [0.3.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.3.0...v0.3.1) (2026-09-24)


### Bug Fixes

* **deps:** bump dotenv from 18.0.1 to 18.0.2 in the production-dependencies group ([#25](https://github.com/chrischall/office-outlook-mcp/issues/25)) ([b35ad69](https://github.com/chrischall/office-outlook-mcp/commit/b35ad698712d6e0e7700ff6dbb7f0de5509fbd82))

## [0.3.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.2.0...v0.3.0) (2026-09-24)


### Features

* confirm writes with a preview token instead of confirm: true ([#23](https://github.com/chrischall/office-outlook-mcp/issues/23)) ([9a083bb](https://github.com/chrischall/office-outlook-mcp/commit/9a083bbd9a72a09ab4fc781f7f9c95359aa4e96d))


### Bug Fixes

* **mail:** frame email and event text as untrusted before it reaches the model ([#21](https://github.com/chrischall/office-outlook-mcp/issues/21)) ([3ff90ef](https://github.com/chrischall/office-outlook-mcp/commit/3ff90efa904f49d86bfe3073aae7a04927f2b674))

## [0.2.0](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.5...v0.2.0) (2026-09-23)


### Features

* **tools:** paginate every list tool via nextLink instead of truncating silently ([#15](https://github.com/chrischall/office-outlook-mcp/issues/15)) ([ba767b2](https://github.com/chrischall/office-outlook-mcp/commit/ba767b2ddc421e7739f372a42ec7083b4387fb51))


### Bug Fixes

* **mail:** clarify that search+unreadOnly is only rejected on a first-page request ([#20](https://github.com/chrischall/office-outlook-mcp/issues/20)) ([7599fbe](https://github.com/chrischall/office-outlook-mcp/commit/7599fbe5e825e17e0c9583c1aff99008bfbe834a))
* **mail:** stop rejecting search+unreadOnly when following a nextLink ([#18](https://github.com/chrischall/office-outlook-mcp/issues/18)) ([2a08a9a](https://github.com/chrischall/office-outlook-mcp/commit/2a08a9a85a1f1091a701ad073a91096b70b8e0a0))

## [0.1.5](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.4...v0.1.5) (2026-09-23)


### Bug Fixes

* **deps:** require zod ^4.6.5 to match @chrischall/mcp-utils 2.4.0 ([#14](https://github.com/chrischall/office-outlook-mcp/issues/14)) ([f443913](https://github.com/chrischall/office-outlook-mcp/commit/f443913a9b8f4bd6fe9feae6907fe83a9bee3d26))
* **deps:** upgrade @chrischall/mcp-utils to 2.4.0 and @fetchproxy/* to 3.2.0 ([#12](https://github.com/chrischall/office-outlook-mcp/issues/12)) ([6556287](https://github.com/chrischall/office-outlook-mcp/commit/6556287dee2f525459766e216347502d0c01da75))

## [0.1.4](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.3...v0.1.4) (2026-09-21)


### Bug Fixes

* allow outlook.cloud.microsoft in hosted egress and the fpx skill profile ([#9](https://github.com/chrischall/office-outlook-mcp/issues/9)) ([50c384b](https://github.com/chrischall/office-outlook-mcp/commit/50c384b24036c13b0793a86b032f8c1755c092ae))

## [0.1.3](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.2...v0.1.3) (2026-09-21)


### Bug Fixes

* **tools:** outlook_send_mail is destructive ([#7](https://github.com/chrischall/office-outlook-mcp/issues/7)) ([17e1df1](https://github.com/chrischall/office-outlook-mcp/commit/17e1df1ccbd1214687fa7bef80e53fdbb56cd2b6))

## [0.1.2](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.1...v0.1.2) (2026-09-21)


### Bug Fixes

* make contacts, the project-scoped launch, capture window and event time zones work live ([#4](https://github.com/chrischall/office-outlook-mcp/issues/4)) ([dda1bb6](https://github.com/chrischall/office-outlook-mcp/commit/dda1bb6793b53124a12c1d5ec237b4c82b0aa17a))


### Documentation

* install in opencode 2 without breaking opencode 1 ([#6](https://github.com/chrischall/office-outlook-mcp/issues/6)) ([79dc7a3](https://github.com/chrischall/office-outlook-mcp/commit/79dc7a30c97fc8de14ba27ed22ac876bd12e9058))

## [0.1.1](https://github.com/chrischall/office-outlook-mcp/compare/v0.1.0...v0.1.1) (2026-09-20)


### Bug Fixes

* **mcpb:** nest the Node floor under compatibility so the bundle packs ([#2](https://github.com/chrischall/office-outlook-mcp/issues/2)) ([36527dc](https://github.com/chrischall/office-outlook-mcp/commit/36527dcac8d92407caf709d431e19084f79e11f3))

## 0.1.0 (2026-09-20)


### Features

* Outlook / Microsoft 365 MCP server ([7a217b7](https://github.com/chrischall/office-outlook-mcp/commit/7a217b7d70c674faf1f8c6bea3d477cd1f53a188))
