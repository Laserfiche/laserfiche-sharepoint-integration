// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

export class SPPermission {
  public static fullMask = new SPPermission();

  constructor(
    public permissions: { High: number; Low: number } = { High: 0, Low: 0 }
  ) { }

  public hasPermission(permission: SPPermission): boolean {
    return true;
  }
}
