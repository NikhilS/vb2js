/*
 * Copyright 2010 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.google.vb2js;

import com.google.common.collect.Sets;

import java.util.Set;
import java.util.Stack;

/**
 * Simple class to keep track of the global state of the VBA code being
 * converted.
 *
 * @author Brian Kernighan
 * @author Nikhil Singhal
 */
final class GlobalState {

  /** Stack of names in With */
  private final Stack<String> withNames;

  /** Names of global vars */
  private final Set<String> globalNames;

  /** Names of local vars */
  private final Set<String> localNames;

  GlobalState() {
    this.withNames = new Stack<String>();

    this.globalNames = Sets.newHashSet();
    this.localNames = Sets.newHashSet();
  }

  void addGlobalName(String name) {
    globalNames.add(name);
  }

  void addLocalName(String name) {
    localNames.add(name);
  }

  /** Add a With name to the stack */
  void addWithName(String name) {
    withNames.push(name);
  }

  void clearLocalNames() {
    localNames.clear();
  }

  /** Get the current With name (ie, the top of the stack) */
  String getWithName() {
    return withNames.peek();
  }

  /** Remove the latest With name from the stack (ie, pop) */
  void popWithName() {
    withNames.pop();
  }

  boolean isArrayName(String name) {
    return (localNames.contains(name) || globalNames.contains(name));
  }
}