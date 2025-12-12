# Workflow: Docker Update & Run

You are an assistant that helps update and restart Docker containers for any project. Follow these steps precisely.

### 1. Check Docker Compose File
- Use `execute_command` to run `ls -la` to confirm you're in the project directory with docker-compose.yml file
- If docker-compose.yml is not found, inform the user and stop the workflow

### 2. Stop and Remove Project Containers
- Use `execute_command` to run `docker-compose down -v` to stop and remove all containers, networks, and volumes for this project
- The `-v` flag ensures complete cleanup of project volumes

### 3. Clean Up Docker System (Optional - Requires User Confirmation)
- **⚠️ WARNING**: Only run `docker system prune -f` if user explicitly confirms
- This command removes ALL unused Docker resources system-wide, which may affect other projects
- Ask user: "Do you want to clean up unused Docker resources? This will remove unused containers, networks, and images from all projects on your system."
- If user agrees, use `execute_command` to run `docker system prune -f`

### 4. Rebuild Images
- Use `execute_command` to run `docker-compose build --no-cache` to force rebuild all images from scratch
- This ensures all code changes are included in the new images

### 5. Start Services with Logs
- Use `execute_command` to run `docker-compose up` to start services and show live logs
- This will start all services defined in docker-compose.yml and display output in real-time

### 6. Monitor Startup
- Watch the output for any error messages or successful startup indicators
- Look for services that depend on others (like frontend depending on backend) to start in correct order

### 7. Verify Services
- After services start, use `execute_command` to run `docker-compose ps` to check all containers are running
- Report the status of each service to the user

### 8. Service URLs (if applicable)
- Common service ports to check:
  - Frontend: typically port 3000, 8080, or 80
  - Backend: typically port 8000, 5000, or 3001
  - Database: typically port 5432, 3306, or 27017

### Prerequisites
- Switch to CODE mode using `switch_mode` tool before starting this workflow, as Docker operations require CODE mode permissions
- Ensure Docker and Docker Compose are installed
- Ensure you're in the project directory containing docker-compose.yml

### Universal Compatibility
This workflow works with any Docker Compose project regardless of:
- Number of services
- Service types (web, database, cache, etc.)
- Port configurations
- Volume configurations
- Network setups

### Troubleshooting Commands
- `docker-compose logs -f [service]` - follow logs for specific service
- `docker-compose exec [service] sh` - shell into running container
- `docker-compose restart [service]` - restart specific service
- `docker-compose down` - stop all services
### Troubleshooting Commands
- `docker-compose logs -f [service]` - follow logs for specific service
- `docker-compose exec [service] sh` - shell into running container
- `docker-compose restart [service]` - restart specific service
- `docker-compose down` - stop all services
- `docker-compose up -d` - start in detached mode if needed

---

## Service Status Report

After completing the workflow, all services should be running. Here's how to check and access them:

### Verify Services Status
```bash
docker-compose ps
```

This command will show all running containers with their status, ports, and service names.

### Service Access URLs
Once services are running, access them at:

| Service Type | Typical URL Pattern | Description |
|--------------|-------------------|-------------|
| **Frontend/Web** | [http://localhost:3000](http://localhost:3000) | Web interface, SPA, or static files |
| **Backend API** | [http://localhost:8000](http://localhost:8000) | REST API endpoints |
| **API Documentation** | [http://localhost:8000/docs](http://localhost:8000/docs) | Swagger/OpenAPI docs (if available) |
| **Admin Panel** | [http://localhost:8080](http://localhost:8080) | Administrative interface |
| **Database** | localhost:5432 | PostgreSQL, MySQL, etc. |
| **Cache/Redis** | localhost:6379 | Redis, Memcached |
| **Message Queue** | localhost:5672 | RabbitMQ, Apache Kafka |

### Monitor and Debug

#### View Real-time Logs
```bash
# All services
docker-compose logs -f

# Specific service
docker-compose logs -f [service-name]

# Last 100 lines
docker-compose logs --tail=100 [service-name]
```

#### Quick Health Check
```bash
# Check container status
docker-compose ps

# Test API endpoints (replace with your actual ports)
curl http://localhost:8000/health
curl http://localhost:3000

# Check if ports are listening
netstat -tlnp | grep :3000
netstat -tlnp | grep :8000
```

#### Common Issues and Solutions
- **Port conflicts**: Check if ports are already in use
- **Container won't start**: Check logs with `docker-compose logs [service]`
- **Network issues**: Restart with `docker-compose down && docker-compose up`
- **Out of memory**: Check Docker Desktop resources allocation

### Next Steps
1. Open your browser to the frontend URL
2. Test API endpoints if applicable
3. Check application logs for any errors
4. Verify database connections if required
- `docker-compose up -d` - start in detached mode if needed
