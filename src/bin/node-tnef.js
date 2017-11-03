#!/usr/bin/env node
import yargs from 'yargs'

export const argv = yargs
    .commandDir('../commands')
    .demandCommand()
    .help()
    .argv