# Changelog

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
