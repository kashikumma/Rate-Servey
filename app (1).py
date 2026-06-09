import streamlit as st
import requests
import pandas as pd
import time
import json
import urllib.request
from math import radians, sin, cos, sqrt, atan2

st.set_page_config(page_title="Metropolis AI", layout="wide")

RADIUS_MILES = 5.0

# --- HAVERSINE DISTANCE ---
def haversine(lat1, lon1, lat2, lon2):
    R = 3958.8
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    return R * 2 * atan2(sqrt(a), sqrt(1 - a))

# --- GEOCODE ADDRESS ---
def geocode_address(address):
    try:
        q = address.replace(" ", "+")
        url = f"https://nominatim.openstreetmap.org/search?q={q}&format=json&limit=1&countrycodes=us"
        req = urllib.request.Request(url, headers={"User-Agent": "MetropolisGarageSearch/1.0"})
        res = urllib.request.urlopen(req, timeout=10)
        data = json.loads(res.read().decode())
        if data:
            return float(data[0]["lat"]), float(data[0]["lon"]), data[0]["display_name"]
    except Exception:
        pass
    return None, None, None

# --- LIVE MAP SEARCH FOR GARAGES (With Backup Server) ---
@st.cache_data(ttl=3600)
def search_live_garages(lat, lon, radius_miles):
    radius_meters = radius_miles * 1609.34
    
    # List of reliable open endpoints to try if one is busy
    endpoints = [
        "https://maps.mail.ru/osm/tools/overpass/api/interpreter",
        "https://overpass-api.de/api/interpreter"
    ]
    
    overpass_query = f"""
    [out:json][timeout:20];
    (
      node["amenity"="parking"](around:{radius_meters},{lat},{lon});
      way["amenity"="parking"](around:{radius_meters},{lat},{lon});
    );
    out center;
    """
    
    data = None
    for url in endpoints:
        try:
            response = requests.post(url, data={'data': overpass_query}, timeout=12)
            if response.status_code == 200:
                data = response.json()
                break  # Success, exit the loop!
        except Exception:
            continue  # Fail silently and try the next server
            
    if not data:
        st.error("The public map servers are temporarily busy. Please wait a moment and hit enter to try again.")
        return []
        
    garages = []
    for element in data.get('elements', []):
        tags = element.get('tags', {})
        g_lat = element.get('lat') or element.get('center', {}).get('lat')
        g_lon = element.get('lon') or element.get('center', {}).get('lon')
        
        if g_lat and g_lon:
            dist = haversine(lat, lon, g_lat, g_lon)
            parking_type = tags.get('parking', 'surface')
            name = tags.get('name', f"Unnamed Parking ({parking_type.title()})")
            
            garages.append({
                "name": name,
                "lat": g_lat,
                "lon": g_lon,
                "distance": dist,
                "type": parking_type.title(),
                "access": tags.get('access', 'Public').title(),
                "operator": tags.get('operator', 'Unknown Operator'),
                "capacity": tags.get('capacity', 'Not Listed')
            })
    
    garages.sort(key=lambda x: x['distance'])
    return garages

# ===================== MAIN UI =====================
st.markdown("# Metropolis Market Intelligence")
st.write("Enter any US address to search live real-world parking garages and lots within **5 miles**.")

search_term = st.text_input("Search a US address", placeholder="e.g. 1455 N Sandburg Terrace, Chicago, IL 60610")

if search_term:
    with st.spinner("Geocoding address..."):
        ref_lat, ref_lon, ref_label = geocode_address(search_term)

    if ref_lat is not None:
        st.markdown(f"**Searching near:** {ref_label}")
        st.caption(f"Coordinates: {ref_lat:.6f}, {ref_lon:.6f}")

        with st.spinner("AI map search locating nearby garages..."):
            found_garages = search_live_garages(ref_lat, ref_lon, RADIUS_MILES)

        if found_garages:
            st.success(f"Found {len(found_garages)} parking facilities within {RADIUS_MILES} miles!")
            
            # Show interactive map
            map_data = pd.DataFrame([
                {"lat": g["lat"], "lon": g["lon"], "name": g["name"]} for g in found_garages
            ])
            st.map(map_data)
            
            # Display garage details
            for g in found_garages:
                st.markdown("---")
                st.markdown(f"### {g['name']} &nbsp; <span style='font-size:14px; color:gray;'>({g['distance']:.2f} mi away)</span>", unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"**Facility Type:** {g['type']}")
                    st.markdown(f"**Operator:** {g['operator']}")
                with col2:
                    st.markdown(f"**Access:** {g['access']}")
                    st.markdown(f"**Capacity:** {g['capacity']} spaces")
        else:
            st.warning(f"No parking garages found within {RADIUS_MILES} miles of this location.")
    else:
        st.error("Could not find that address. Please check your spelling or try including city/state.")
