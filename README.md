# The Cayman theme

## Usage

To use the Cayman theme:

    Add the following to your site's _config.yml:

    remote_theme: pages-themes/cayman@v0.2.0
    plugins:
    - jekyll-remote-theme # add this line to the plugins list if you already have one

    Optionally, if you'd like to preview your site on your computer, add the following to your site's Gemfile:

    gem "github-pages", group: :jekyll_plugins

Customizing
Configuration variables

Cayman will respect the following variables, if set in your site's _config.yml:

title: [The title of your site]
description: [A short description of your site's purpose]

Additionally, you may choose to set the following optional variables:

show_downloads: ["true" or "false" (unquoted) to indicate whether to provide a download URL]
google_analytics: [Your Google Analytics tracking ID]

Stylesheet

If you'd like to add your own custom styles:

    Create a file called /assets/css/style.scss in your site
    Add the following content to the top of the file, exactly as shown: 
