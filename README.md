# PnP Modern Search - Extensibility samples

This project centralizes extensibility samples made by the community for the [PnP Modern Search solution (v4)](https://github.com/microsoft-search/pnp-modern-search). By "extensions", we mean:

- Custom web components
- Custom search box suggestions providers
- Custom Handlebars helpers
- Custom layouts
- Custom data sources
- Custom adaptive cards actions
- Custom query modifiers

## Get Started

- Install and configure the [PnP Modern Search solution](https://microsoft-search.github.io/pnp-modern-search/installation/) in your SharePoint Online environment.
- Download the latest release of the [PnP Modern Search extensibility samples](https://github.com/microsoft-search/pnp-modern-search-extensibility-samples/releases) and deploy it to your tenant or site collection app catalog.

## Build Environment Setup

 1. Set the project to the desired environemnt (`base`, `dev`, `uat`, or `prod`)

     `npm run set-env:prod`

 2. Bundle and package solution

     `gulp clean --ship`
     `gulp bundle --ship`
     `gulp package-solution --ship`

 3. Reset the project to the base environment

    `npm run set-env:base`

## Adding extension samples

- Clone this repository
- Add your samples to the existing `search-extensibility-samples` SPFx project following the [step-by-step guide](https://microsoft-search.github.io/pnp-modern-search/extensibility/).
- Create documentation for your extensions (procedure to be determined)
- Submit your PR

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

node v18.18.0
