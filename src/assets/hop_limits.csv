platform,max_hops,description,notes
windows_x86,4,Windows (Intel/AMD),xHCI supports 5 hubs per spec; practical limits often around 4 external hubs due to endpoint constraints. Recommended is powered hubs for multiple devices.
windows_arm,4,Windows (ARM),ARM PCs follow standard xHCI limits but may have specific controller limitations depending on hardware vendor.
mac_intel,7,Mac (Intel),Intel Macs support deeper chains (7 tiers). Use powered hubs or Thunderbolt docks for multi-device setups.
mac_apple_silicon,3,Mac (Apple Silicon),Built-in hub per USB-C port consumes 1 tier. Practical limit: built-in + 2 external hubs + device. Less is more unfortunate due to Apple Silicon implementation.
linux_x86,4,Linux (Intel/AMD),Linux follows xHCI spec but is limited by kernel endpoint allocation.
linux_arm,4,Linux (ARM),ARM systems (e.g. Raspberry Pi) may have tighter limits; specific controller limitations may apply.
iphone_lightning,2,iPhone (Lightning),iPhone Lightning devices have limited hub support; avoid deep hub chains for reliable operation. Avoid if possible.
iphone_usbc,3,iPhone (USB-C),Less is more unfortunate due to Apple Silicon implementation. Use powered hubs if multiple devices needed.
samsung_usbc,4,Samsung (USB-C),Practical hub limits vary; specific chipset (Exynos/Qualcomm) and vendor firmware may restrict deep chains.
android_usbc,4,Android (USB-C),Practical hub limits vary; specific chipset (Qualcomm/MediaTek/Tensor) and vendor firmware may restrict deep chains.
ipad_lightning,2,iPad (Lightning),iPad Lightning devices have limited hub support; avoid deep chains. Avoid if possible.
ipad_usbc,3,iPad (USB-C),M-series supports limited chains; A-series has tighter limits. Use powered hubs or compatible docks for multi-device setups.
