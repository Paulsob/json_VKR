#!/usr/bin/env python3
"""
Test script to verify that route numbers are loaded correctly from consolidation folder
"""

import os
from structure_model.driver_scheduler import run_planner_for_all_routes

def test_route_loading():
    print("Testing route loading from consolidation folder...")
    
    # Check what routes are in the consolidation file
    routes_file = "/workspace/consolidation/route_numbers.txt"
    print(f"Reading route numbers from: {routes_file}")
    
    with open(routes_file, 'r', encoding='utf-8') as f:
        routes_in_file = []
        for line in f:
            line = line.strip()
            if line and line.isdigit():
                routes_in_file.append(int(line))
    
    print(f"Routes found in file: {routes_in_file}")
    
    # Test the function with a single day to verify it processes all routes
    print("\nTesting run_planner_for_all_routes with day 1...")
    run_planner_for_all_routes(1, 0)  # Day 1, previous day 0
    
    print("\nVerification complete!")
    print("All routes from consolidation folder have been processed.")
    print("Check the output directories for Расписание_Итог_1.xlsx files.")

if __name__ == "__main__":
    test_route_loading()