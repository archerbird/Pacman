# Pacman
An OOP implementation of Pacman in Excel

This is very much a WIP. Some explaination of the code and architecture can be found [here](https://codereview.stackexchange.com/q/248785/206696) for now.

Right now, until I've got a fully working game, the develop branch is default. 

## Get Started
You should be able to simply clone and launch the Pacman.xlsm. Sheet1 has a "hardcoded" maze. Eventually, I'd like to teach excel how to draw a maze from scratch. 
You can start a game by running Client.Prototype(). As of now, the only things that work are:
1. You can move pacman by using the arrow keys
1. Ghosts will navigate themselves around the maze

For now, you can't die, but you also can't eat! 

## Rubberduck
Its really helpful to have the Rubberduck VBE Add-in installed. You will be able to see a nicer folder structure for the project.

## Disclaimers
* I have only ever run this on my machine. I don't know how well things like resolution and game ticking speed will work on your machine. 
* It definitely won't work on macOS
* I have experienced some really nasty crashes while trying new things in development. Crashes that completely broke the vbproject.bin. Here are some things that seem to cause crash issues:
  1. Using custom collections with enumerators
  1. Having the MapManager load its own map file. (if the controller does this, it seems to not crash???)
  
  
