# vmware-baseline
A VMWare baseline compliance script and modules based around the [VMWare Hardening Guide](https://www.vmware.com/security/hardening-guides)

This is '''NOT''' intended to be a Daily Report. [vCheck](https://github.com/alanrenouf/vCheck-vSphere) satisfies that much better. This is intended to be used part and parcel with compliance scanners like Nessus.

Originally written against VMWare 5.5U2. Initial tests indicate works with PowerCLI 6 but have not had opportunity to test against 6.0.

This script is quite old and could be vastly improved. However, I do think the original approach of serializing state as hashtables and merging them was correct. Nonetheless, this script has been used to meet various controls and POA&Ms.

## Design Goals Decisions
- Provide declarative INI-style "desried state" configurations (although some aspects had to ultimately be hardcoded)
- Provide both interactive and automatic modes
- Generate reports (HTML)
- Ability to merge multiple levels of the VMWare Baseline
- Attempted to create "rollback" files, though this proved difficult has the VC API does not provide easy ways to `NULL` or reset a certain values. Ergo, some assumptions had to be made about boolean values and such.

 