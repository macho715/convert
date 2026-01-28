"""
Professional Stakeholder Network Visualization for HVDC AGI Transformer Transportation Project
Built with modern design principles: MSA, UX enhancement, clean code, DDD
"""

import json
import re
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Tuple, Any

import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import FancyBboxPatch
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ============================================================================
# Configuration: Professional Styling (2024 Design Standards)
# ============================================================================

COLORS = {
    'samsung': {
        'primary': '#003366',
        'light': '#E3F2FD',
        'gradient': ['#003366', '#1976D2']
    },
    'aries': {
        'primary': '#008B8B',
        'light': '#E0F7FA',
        'gradient': ['#008B8B', '#00ACC1']
    },
    'adnoc': {
        'primary': '#2E7D32',
        'light': '#E8F5E9',
        'gradient': ['#2E7D32', '#66BB6A']
    },
    'mammoet': {
        'primary': '#D84315',
        'light': '#FBE9E7',
        'gradient': ['#D84315', '#FF7043']
    },
    'lct_bushra': {
        'primary': '#546E7A',
        'light': '#ECEFF1',
        'gradient': ['#546E7A', '#78909C']
    }
}

# Organizational mapping
ORG_MAPPING = {
    'Samsung C&T': 'samsung',
    'Aries Marine': 'aries',
    'ADNOC L&S': 'adnoc',
    'Mammoet': 'mammoet',
    'LCT Bushra': 'lct_bushra',
    'Khalid Faraj Shipping': 'lct_bushra'
}

# Milestones for decision flow
MILESTONES = [
    ('2025-09-19', 'Project Initiation', 'Rocket icon'),
    ('2025-09-22', 'Scope Proposal', 'Document icon'),
    ('2025-10-10', 'Load Calculation', 'Calculator icon'),
    ('2025-10-21', 'Role Clarification', 'Handshake icon'),
    ('2025-10-25', 'Bow Deck Approval', 'Check shield icon'),
    ('2025-10-26', 'ELC Data Provided', 'Database icon'),
    ('2025-10-28', 'Final Approval', 'Trophy icon')
]

# ============================================================================
# Phase 1: Data Parsing (Clean Architecture)
# ============================================================================

def parse_json_ld(filepath: str) -> Dict[str, Any]:
    """Parse JSON-LD from machine-readable email file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract JSON-LD block
    json_start = content.find('```json')
    json_end = content.find('```', json_start + 7)
    
    if json_start == -1 or json_end == -1:
        raise ValueError("JSON-LD block not found")
    
    json_str = content[json_start + 7:json_end].strip()
    return json.loads(json_str)

def extract_participants(data: Dict[str, Any]) -> pd.DataFrame:
    """Extract participants into a structured DataFrame."""
    participants = []
    for p in data['participants']:
        org_key = ORG_MAPPING.get(p['org'], 'samsung')
        participants.append({
            'name': p['name'],
            'email': p.get('email', ''),
            'org': p['org'],
            'org_key': org_key,
            'role': p.get('role', '')
        })
    
    return pd.DataFrame(participants)

def extract_messages(data: Dict[str, Any]) -> pd.DataFrame:
    """Extract messages into a structured DataFrame."""
    messages = []
    for msg in data['messages']:
        # Extract from email
        from_match = re.search(r'<([^>]+)>', msg['from'])
        from_email = from_match.group(1) if from_match else msg['from']
        
        # Extract to emails
        to_emails = []
        if 'to' in msg:
            for to in msg['to']:
                to_match = re.search(r'<([^>]+)>', to)
                to_emails.append(to_match.group(1) if to_match else to)
        
        messages.append({
            'id': msg['id'],
            'order': msg.get('order', 0),
            'date': msg.get('isoDate', ''),
            'from': from_email,
            'to': to_emails if to_emails else [],
            'subject': msg.get('subject', ''),
            'summary': msg.get('summary', ''),
            'topics': data.get('topics', [])
        })
    
    return pd.DataFrame(messages)

def extract_org_message_counts(participants_df: pd.DataFrame, messages_df: pd.DataFrame) -> Dict[str, int]:
    """Count messages per organization."""
    org_counts = defaultdict(int)
    
    # Create name to org mapping
    name_to_org = dict(zip(participants_df['name'], participants_df['org']))
    email_to_org = dict(zip(participants_df['email'], participants_df['org']))
    
    for _, msg in messages_df.iterrows():
        # Count from org
        from_org = email_to_org.get(msg['from'])
        if from_org:
            org_counts[from_org] += 1
        
        # Count to orgs
        for to in msg['to']:
            to_org = email_to_org.get(to)
            if to_org:
                org_counts[to_org] += 1
    
    return dict(org_counts)

# ============================================================================
# Phase 2: Graph Construction (Factory Pattern)
# ============================================================================

def build_organizational_network(participants_df: pd.DataFrame, messages_df: pd.DataFrame) -> nx.DiGraph:
    """Build organizational-level network graph."""
    G = nx.DiGraph()
    
    # Get unique organizations
    orgs = participants_df['org'].unique()
    org_key_map = dict(zip(participants_df['org'], participants_df['org_key']))
    
    # Add organizations as nodes
    org_message_counts = extract_org_message_counts(participants_df, messages_df)
    
    for org in orgs:
        org_key = org_key_map[org]
        participant_count = len(participants_df[participants_df['org'] == org])
        message_count = org_message_counts.get(org, 0)
        
        G.add_node(org, 
                   org_key=org_key,
                   participant_count=participant_count,
                   message_count=message_count,
                   size=message_count * 50 + participant_count * 100)
    
    # Add edges based on message exchanges
    email_to_org = dict(zip(participants_df['email'], participants_df['org']))
    
    org_message_map = defaultdict(lambda: defaultdict(int))
    
    for _, msg in messages_df.iterrows():
        from_org = email_to_org.get(msg['from'])
        for to_email in msg['to']:
            to_org = email_to_org.get(to_email)
            if from_org and to_org and from_org != to_org:
                org_message_map[from_org][to_org] += 1
    
    # Add weighted edges
    for from_org, targets in org_message_map.items():
        for to_org, count in targets.items():
            G.add_edge(from_org, to_org, weight=count)
    
    return G

def build_individual_network(participants_df: pd.DataFrame, messages_df: pd.DataFrame) -> nx.DiGraph:
    """Build individual participant network graph."""
    G = nx.DiGraph()
    
    email_to_person = {}
    person_to_org = {}
    
    for _, p in participants_df.iterrows():
        email_to_person[p['email']] = p['name']
        person_to_org[p['name']] = p['org_key']
    
    # Add participants as nodes
    for _, p in participants_df.iterrows():
        G.add_node(p['name'],
                   org=p['org'],
                   org_key=p['org_key'],
                   email=p['email'],
                   role=p.get('role', ''))
    
    # Count message exchanges
    person_message_map = defaultdict(lambda: defaultdict(int))
    
    for _, msg in messages_df.iterrows():
        from_person = email_to_person.get(msg['from'])
        for to_email in msg['to']:
            to_person = email_to_person.get(to_email)
            if from_person and to_person and from_person != to_person:
                person_message_map[from_person][to_person] += 1
    
    # Add weighted edges
    for from_person, targets in person_message_map.items():
        for to_person, count in targets.items():
            G.add_edge(from_person, to_person, weight=count)
    
    return G

def build_decision_flow_network() -> nx.DiGraph:
    """Build decision flow timeline network."""
    G = nx.DiGraph()
    
    prev_date = None
    prev_milestone = None
    
    for date_str, name, icon in MILESTONES:
        G.add_node(name, date=date_str, icon=icon)
        
        if prev_milestone:
            # Calculate duration
            date = datetime.strptime(date_str, '%Y-%m-%d')
            prev_dt = datetime.strptime(prev_date, '%Y-%m-%d')
            duration = (date - prev_dt).days
            
            G.add_edge(prev_milestone, name, duration=duration)
        
        prev_date = date_str
        prev_milestone = name
    
    return G

# ============================================================================
# Phase 3: Layout & Styling
# ============================================================================

def apply_organizational_layout(G: nx.DiGraph) -> Dict[str, Tuple[float, float]]:
    """Apply force-directed layout to organizational network."""
    pos = nx.spring_layout(G, k=2, iterations=50, seed=42)
    return pos

def apply_individual_layout(G: nx.DiGraph) -> Dict[str, Tuple[float, float]]:
    """Apply hierarchical layout to individual network."""
    pos = nx.spring_layout(G, k=1.5, iterations=50, seed=42)
    return pos

def apply_timeline_layout(G: nx.DiGraph) -> Dict[str, Tuple[float, float]]:
    """Apply linear timeline layout to decision flow."""
    pos = {}
    nodes = list(G.nodes())
    
    for i, node in enumerate(nodes):
        pos[node] = (i * 2, 0)
    
    return pos

def create_organizational_visualization(G: nx.DiGraph, output_path: str):
    """Create professional organizational network visualization."""
    pos = apply_organizational_layout(G)
    
    fig, ax = plt.subplots(figsize=(16, 12))
    
    # Draw edges
    for edge in G.edges(data=True):
        u, v, data = edge
        weight = data.get('weight', 1)
        x1, y1 = pos[u]
        x2, y2 = pos[v]
        
        color = 'gray'
        width = 1 + weight * 0.5
        
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', color=color, 
                                 lw=width, alpha=0.6))
    
    # Draw nodes
    for node, data in G.nodes(data=True):
        org_key = data.get('org_key', 'samsung')
        size = data.get('size', 1000)
        x, y = pos[node]
        
        color = COLORS[org_key]['primary']
        light = COLORS[org_key]['light']
        
        # Node
        circle = plt.Circle((x, y), size * 0.001, color=color, zorder=3)
        ax.add_patch(circle)
        
        # Label
        ax.text(x, y + size * 0.001 + 0.05, node, 
               ha='center', va='bottom', fontsize=10, weight='bold')
    
    ax.set_xlim(-2, 2)
    ax.set_ylim(-1.5, 1.5)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title('HVDC Project: Organizational Communication Network', 
                fontsize=16, weight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"✓ Created organizational network: {output_path}")

def create_individual_visualization(G: nx.DiGraph, output_path: str):
    """Create professional individual network visualization."""
    pos = apply_individual_layout(G)
    
    fig, ax = plt.subplots(figsize=(20, 16))
    
    # Group by organization for color coding
    org_groups = {}
    for node, data in G.nodes(data=True):
        org_key = data.get('org_key', 'samsung')
        if org_key not in org_groups:
            org_groups[org_key] = []
        org_groups[org_key].append(node)
    
    # Draw edges
    for edge in G.edges(data=True):
        u, v, data = edge
        weight = data.get('weight', 1)
        x1, y1 = pos[u]
        x2, y2 = pos[v]
        
        color = 'lightgray'
        width = 0.5 + weight * 0.3
        
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', color=color, 
                                 lw=width, alpha=0.4))
    
    # Draw nodes by organization
    for org_key, nodes in org_groups.items():
        color = COLORS[org_key]['primary']
        
        for node in nodes:
            if node in pos:
                x, y = pos[node]
                circle = plt.Circle((x, y), 0.08, color=color, zorder=3)
                ax.add_patch(circle)
                
                # Label
                ax.text(x, y + 0.12, node.split()[0], 
                       ha='center', va='bottom', fontsize=8)
    
    ax.set_xlim(-3, 3)
    ax.set_ylim(-2, 2)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title('HVDC Project: Individual Communication Network', 
                fontsize=16, weight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"✓ Created individual network: {output_path}")

def create_timeline_visualization(G: nx.DiGraph, output_path: str):
    """Create professional decision timeline visualization."""
    pos = apply_timeline_layout(G)
    
    fig, ax = plt.subplots(figsize=(24, 8))
    
    # Draw nodes as milestone cards
    for node, (x, y) in pos.items():
        data = G.nodes[node]
        date = data.get('date', '')
        
        # Card background
        rect = FancyBboxPatch((x-0.8, y-0.4), 1.6, 0.8, 
                             boxstyle="round,pad=0.1",
                             facecolor='white', edgecolor='#003366',
                             linewidth=2, zorder=2)
        ax.add_patch(rect)
        
        # Date label
        ax.text(x, y+0.3, date[5:], ha='center', va='center',
               fontsize=8, weight='bold')
        
        # Milestone name
        ax.text(x, y, node, ha='center', va='center',
               fontsize=9, wrap=True)
    
    # Draw edges
    for edge in G.edges(data=True):
        u, v, data = edge
        x1, y1 = pos[u]
        x2, y2 = pos[v]
        
        duration = data.get('duration', 0)
        color = '#FF7043' if duration > 8 else '#66BB6A'
        width = 3 if duration > 8 else 2
        
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                   arrowprops=dict(arrowstyle='->', color=color, 
                                 lw=width))
    
    ax.set_xlim(-1, len(G.nodes()) * 2 - 1)
    ax.set_ylim(-1, 1)
    ax.axis('off')
    ax.set_title('HVDC Project: Decision Flow Timeline (Sep 19 - Oct 28, 2025)', 
                fontsize=16, weight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"✓ Created timeline network: {output_path}")

# ============================================================================
# Main Execution
# ============================================================================

def main():
    """Main execution function."""
    print("=" * 80)
    print("HVDC Stakeholder Network Visualization")
    print("Professional Grade - Modern Design Principles")
    print("=" * 80)
    
    # Parse data
    print("\n[Phase 1] Parsing JSON-LD data...")
    data = parse_json_ld('mrconvert_v1/AGI Transformers Transportation_email.md')
    participants_df = extract_participants(data)
    messages_df = extract_messages(data)
    
    print(f"  ✓ Loaded {len(participants_df)} participants")
    print(f"  ✓ Loaded {len(messages_df)} messages")
    
    # Build networks
    print("\n[Phase 2] Building network graphs...")
    org_network = build_organizational_network(participants_df, messages_df)
    individual_network = build_individual_network(participants_df, messages_df)
    timeline_network = build_decision_flow_network()
    
    print(f"  ✓ Organization network: {org_network.number_of_nodes()} nodes, {org_network.number_of_edges()} edges")
    print(f"  ✓ Individual network: {individual_network.number_of_nodes()} nodes, {individual_network.number_of_edges()} edges")
    print(f"  ✓ Timeline network: {timeline_network.number_of_nodes()} nodes, {timeline_network.number_of_edges()} edges")
    
    # Generate visualizations
    print("\n[Phase 3] Generating visualizations...")
    create_organizational_visualization(org_network, 'stakeholder_network_organizational.png')
    create_individual_visualization(individual_network, 'stakeholder_network_individuals.png')
    create_timeline_visualization(timeline_network, 'decision_flow_timeline.png')
    
    print("\n" + "=" * 80)
    print("✓ All visualizations created successfully!")
    print("=" * 80)

if __name__ == '__main__':
    main()

