# Papertrail

![Project logo](https://github.com/NASA-JPL-Teamtools-Studio/teamtools-documentation/blob/main/docs/images/tts_image_artifacts/papertrail.png)

## About Teamtools Studio

Teamtools Studio Utilities is part of JPL's Teamtools Studio (TTS).

TTS is an effort originated in JPL's Planning and Execution section to centralize shared repositories across missions. This benefits JPL by reducing cost through reducing duplicated code, collaborating across missions, and unifying standards for development and design across JPL.

Although Planning and Execution is primarily concerned with flight operations, the TTS suite has been generalized and atomized to the point where many of these tools are applicable during other mission phases and even in non-spaceflight contexts. Through our work flying space missions, we hope to provide tools to the open source community that have utility in data analysis or planning for any complex system where failure is not an option.

For more infomation on how to contribute, and how these libraries form a complete ecosystem for high reliability data analysis, see the [Full TTS Documentation](https://nasa-jpl-teamtools-studio.github.io/teamtools_documentation/).

## What is Papertrail?

### Overview

Papertrail is a unified reporting library for Teamtools Studio.

Communicating complex information succinctly is a key part of Systems Engineering, but the ease of that communication
is often hampered by the fact that engineers who are deep experts in their specific disciplines often do not have the
time or interest to learn the nuances of each new reporting tool made available to them. Often solutions are thrown
together quickly and made to be "good enough". Furthermore it is a common pattern that these "good enough" solutions
are often completed in isolation, with each subteam on a project doing very similar work, but none of that work being
similar _enough_ for other teams to leverage it

Papertrail is meant to replace that "good enough" work with a library that is genuinely good. Instead of each team
writing their own reporting tools, the Teamtools Studio has written Papertrail to be a more advanced way to interact
with each reporting tool and other agnostic file formats that projects may use including:

* Cacher (JPL-internal reporting tool)
* Confluence
* Excel Files
* HTML Files
* Microsoft Word

Papertrail also provides a common logging interface for all projects to use so log messages can be interoperable

### Projects Currently Supported

* Europa Clipper
* Mars 2020/Perseverance
* Mars Science Laboratory/Curiosity
* Mars Reconnaisance Orbiter
* NISAR
* Orbiting Carbon Observatory 2
* Sample Retreival Lander

## Architecture

### TTS dependencies

* TTS Utilities
* HTML Utilities
