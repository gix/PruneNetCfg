# PruneNetCfg

A small utility to list and deinstall Windows network devices.

The primary reason for this utility is to remove zombied VirtualBox host-only adapters.
These adapters do not show up anywhere in the VirtualBox or Windows network UI. Newly created adapters are then named
`VirtualBox Host-Only Ethernet Adapter #2` and so forth.


## Usage

    PruneNetCfg [-d] [<filter>]

    -d         Prompts for each device whether it should be deinstalled
    <filter>   Display name filter


## Examples

List network devices:

    PruneNetCfg

List network devices filtered by display name:

    PruneNetCfg "VirtualBox Host-Only Ethernet Adapter"

Select and deinstall network devices:

    PruneNetCfg -d
    PruneNetCfg -d "VirtualBox Host-Only Ethernet Adapter"
